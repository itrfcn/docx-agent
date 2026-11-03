"""
向量库模块
----------
封装 Embedding + Chroma 的初始化、入库、检索、清空逻辑。
向量数据持久化到本地文件夹，重启不丢失。
"""

from __future__ import annotations

import shutil
from pathlib import Path

# HF 离线模式等环境变量统一在 utils/__init__.py 设置（须先于 huggingface_hub 导入）

from langchain_core.documents import Document
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_chroma import Chroma
from langchain_huggingface import HuggingFaceEmbeddings

from utils.config import cfg

# Chroma collection 名称（业务隔离）
COLLECTION_NAME = "enterprise_docs"

# bge 官方约定：查询向量需加检索指令前缀（文档侧不加），可提升召回率
QUERY_INSTRUCTION = "为这个句子生成表示以用于检索相关文章："

# 全局单例向量库实例（延迟初始化）
_vectorstore: Chroma | None = None
_embedding: HuggingFaceEmbeddings | None = None


def get_embedding() -> HuggingFaceEmbeddings:
    """本地 HuggingFace Embedding 单例。"""
    global _embedding
    if _embedding is None:
        # bge-small-zh-v1.5 推荐使用规范化（normalize_embeddings）
        _embedding = HuggingFaceEmbeddings(
            model_name=cfg.embedding_model,
            model_kwargs={"device": "cpu"},
            encode_kwargs={"normalize_embeddings": True},
        )
    return _embedding


def get_vectorstore() -> Chroma:
    """带持久化的 Chroma 单例。"""
    global _vectorstore
    if _vectorstore is None:
        cfg.ensure_dirs()
        _vectorstore = Chroma(
            collection_name=COLLECTION_NAME,
            embedding_function=get_embedding(),
            persist_directory=str(cfg.chroma_persist_dir),
        )
    return _vectorstore


def _make_splitter() -> RecursiveCharacterTextSplitter:
    """构造文本切片器（按中文字符递归切分）。"""
    return RecursiveCharacterTextSplitter(
        chunk_size=cfg.chunk_size,
        chunk_overlap=cfg.chunk_overlap,
        separators=["\n\n", "\n", "。", "！", "？", "；", "，", " ", ""],
    )


def ingest_document(text: str, filename: str) -> int:
    """
    将单个文档文本切片后入库。

    Args:
        text:     文档完整文本
        filename: 文档原始文件名（存为 metadata，用于引用）

    Returns:
        成功入库的切片数量
    """
    splitter = _make_splitter()
    chunks = splitter.split_text(text)

    docs = [
        Document(
            page_content=chunk,
            metadata={"source": filename},
        )
        for chunk in chunks
    ]

    if not docs:
        return 0

    vs = get_vectorstore()
    vs.add_documents(docs)
    return len(docs)


def list_sources() -> list[str]:
    """
    列出知识库中所有已入库的**唯一**文档名称。
    """
    vs = get_vectorstore()
    data = vs.get(include=["metadatas"])
    metadatas = data.get("metadatas") or []
    sources = sorted({m.get("source", "") for m in metadatas if m.get("source")})
    return sources


def _try_remove_persist_files() -> bool:
    """
    兜底：尝试物理删除持久化目录内的文件（Windows 下常被进程占用）。

    Returns:
        True 表示仍有文件残留（被占用或删除失败）。
    """
    residual = False
    persist_dir: Path = cfg.chroma_persist_dir
    if not persist_dir.exists():
        return False

    for child in persist_dir.iterdir():
        try:
            if child.is_file() or child.is_symlink():
                child.unlink(missing_ok=True)
            elif child.is_dir():
                shutil.rmtree(child, ignore_errors=True)
        except OSError:
            pass
        # 删除后仍存在说明被占用
        if child.exists():
            residual = True
    return residual


def clear_knowledge_base() -> dict:
    """
    清空整个知识库。

    优先在 collection 层删除数据（delete_collection + 重建同名空集合），
    不直接物理删除正在被进程占用的 sqlite 文件，避免 Windows WinError 32。

    Returns:
        {
            "cleared": bool,          # 数据是否已清空
            "documents": list[str],  # 清空前的文档清单（manifest 列表）
            "leftover": bool,        # 是否残留历史文件（不影响使用）
        }
    """
    global _vectorstore

    # 1. 记录清空前的文档清单，供界面提示
    try:
        documents = list_sources()
    except Exception:
        documents = []

    cleared = False
    vs: Chroma | None = None
    try:
        vs = _vectorstore or get_vectorstore()
    except Exception:
        vs = None

    if vs is not None:
        # 2. 直接删除整个 collection（数据层面清空，不触碰被占用的文件）
        try:
            vs.delete_collection()
            _vectorstore = None
            cleared = True
        except Exception:
            # 3. 降级方案：删除 collection 内的全部记录
            try:
                ids = (vs.get() or {}).get("ids") or []
                if ids:
                    vs.delete(ids=ids)
                _vectorstore = vs
                cleared = True
            except Exception:
                cleared = False

    # 4. 重建同名空 collection，保证后续上传 / 检索立即可用（无需重启）
    if cleared:
        try:
            _vectorstore = Chroma(
                collection_name=COLLECTION_NAME,
                embedding_function=get_embedding(),
                persist_directory=str(cfg.chroma_persist_dir),
            )
        except Exception:
            _vectorstore = None

    # 5. 成功清空后，尽力删除无引用的孤儿文件（如已废弃的 HNSW 段索引）。
    #    chroma.sqlite3 通常被运行中的进程锁定，跳过（其中已无业务数据，不影响正确性）。
    leftover = False
    if cleared:
        for child in cfg.chroma_persist_dir.glob("*"):
            if child.name.startswith("chroma.sqlite3"):
                continue
            try:
                if child.is_dir():
                    shutil.rmtree(child)
                else:
                    child.unlink()
            except OSError:
                leftover = True

    # 6. 仅在 collection 清理失败时兜底尝试删文件，且容错处理占用异常
    if not cleared:
        leftover = _try_remove_persist_files()
        if not leftover:
            _vectorstore = None
            cleared = True
        # 文件被占用时删除失败，等价于没有产生新的残留
        leftover = bool(leftover)

    return {"cleared": cleared, "documents": documents, "leftover": leftover}


def similarity_search(query: str, top_k: int | None = None) -> list[Document]:
    """
    根据 query 做相似度检索，返回 LangChain Document 列表。
    查询侧附加 bge 检索指令前缀（文档入库时不含前缀）。
    """
    k = top_k or cfg.retriever_top_k
    vs = get_vectorstore()
    return vs.similarity_search(f"{QUERY_INSTRUCTION}{query}", k=k)
