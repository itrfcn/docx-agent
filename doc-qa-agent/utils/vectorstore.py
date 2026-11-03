"""
向量库模块
----------
封装 Embedding + Chroma 的初始化、入库、检索、清空逻辑。
向量数据持久化到本地文件夹，重启不丢失。
"""

from __future__ import annotations

import hashlib
import json
import shutil
from datetime import datetime
from pathlib import Path

# HF 离线模式等环境变量统一在 utils/__init__.py 设置（须先于 huggingface_hub 导入）

from langchain_core.documents import Document
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_chroma import Chroma
from langchain_huggingface import HuggingFaceEmbeddings

from utils.config import cfg

# Chroma collection 名称（业务隔离）
COLLECTION_NAME = "enterprise_docs"

# 距离空间：cosine。使用 cosine 后 similarity_search_with_relevance_scores
# 返回的分数才是可直接用于阈值过滤的「1 - 余弦距离」。
# 注意：Chroma 的 collection metadata 在创建后不可修改，切换空间需重建向量库。
COLLECTION_METADATA = {"hnsw:space": "cosine"}

# 入库清单（manifest）：记录每个 source 的 content_hash 与 chunk_count，
# 用于判断文件内容是否变化、是否需要重建。与向量库目录同级。
MANIFEST_PATH = cfg.chroma_persist_dir.parent / "manifest.json"

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
            collection_metadata=COLLECTION_METADATA,
        )
    return _vectorstore


def _load_manifest() -> dict[str, dict]:
    """读取 manifest 的 documents 字段；文件缺失或损坏时返回空字典。"""
    try:
        with MANIFEST_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
        docs = data.get("documents") if isinstance(data, dict) else None
        return docs if isinstance(docs, dict) else {}
    except (OSError, ValueError):
        return {}


def _save_manifest(documents: dict[str, dict]) -> None:
    """原子写入 manifest（失败不阻断入库）。"""
    try:
        MANIFEST_PATH.parent.mkdir(parents=True, exist_ok=True)
        tmp = MANIFEST_PATH.with_suffix(".tmp")
        with tmp.open("w", encoding="utf-8") as f:
            json.dump({"version": 1, "documents": documents}, f, ensure_ascii=False, indent=2)
        tmp.replace(MANIFEST_PATH)
    except OSError:
        pass


def content_hash(text: str) -> str:
    """计算文档内容哈希（用于判断内容是否变化）。"""
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def get_manifest_entry(filename: str) -> dict | None:
    """获取单个文档的 manifest 记录（不存在返回 None）。"""
    return _load_manifest().get(filename)


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
    # 先按 source 删除旧数据，再写入新数据：
    # 同一文件重复入库不会产生重复片段，文件内容变化后也不会残留旧内容。
    delete_document(filename)
    vs.add_documents(docs)

    # 记录内容哈希与切片数，供导入脚本判断是否需要重建
    manifest = _load_manifest()
    manifest[filename] = {
        "content_hash": content_hash(text),
        "chunk_count": len(docs),
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    _save_manifest(manifest)
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

    # 1.5 快照「清理前已存在的 segment 目录」。
    #     重建 collection 后会生成新的 segment 目录，只删快照里的旧目录，
    #     避免把刚建好的新集合目录一起删掉（那会导致清空后首次写入异常）。
    persist_dir: Path = cfg.chroma_persist_dir
    old_dirs: list[Path] = [p for p in persist_dir.iterdir() if p.is_dir()] if persist_dir.exists() else []

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
                collection_metadata=COLLECTION_METADATA,
            )
        except Exception:
            _vectorstore = None
        # 清空 manifest，避免残留已删除文档的哈希记录
        _save_manifest({})

    # 5. 成功清空后，仅删除「快照里的旧 segment 目录」（新目录不在快照中，不会被误删）。
    #    chroma.sqlite3 被运行中的进程锁定，保持不动（其中已无业务数据，不影响正确性）。
    #    文件被占用时容错跳过，只标记 leftover，不抛 WinError 32。
    leftover = False
    if cleared:
        for child in old_dirs:
            try:
                if child.exists():
                    shutil.rmtree(child)
            except OSError:
                pass
            if child.exists():
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


def delete_document(filename: str) -> int:
    """
    删除知识库中指定文档的全部切片（按 metadata.source 精确匹配）。

    Args:
        filename: 文档名（与 list_sources 返回值一致）

    Returns:
        实际删除的切片数量
    """
    vs = get_vectorstore()
    data = vs.get(where={"source": filename})
    ids = (data or {}).get("ids") or []
    if ids:
        vs.delete(ids=ids)

    # 同步清理 manifest 记录
    manifest = _load_manifest()
    if filename in manifest:
        del manifest[filename]
        _save_manifest(manifest)
    return len(ids)


def similarity_search(query: str, top_k: int | None = None) -> list[Document]:
    """
    根据 query 做相似度检索，返回 LangChain Document 列表。

    - 查询侧附加 bge 检索指令前缀（文档入库时不含前缀）
    - 使用 similarity_search_with_relevance_scores 取带分数结果（要求 cosine 空间）
    - 低于 cfg.relevance_threshold 的片段直接丢弃；全部低于阈值时返回空列表，
      由调用方明确告知「未检索到相关片段」，避免硬凑无关内容
    - 按**来源**去重：每个文档最多保留 1 个片段（过量取 2 倍候选再过滤），
      避免 top_k 名额被同一文件占用，保证返回结果覆盖更多文档
    """
    k = top_k or cfg.retriever_top_k
    threshold = cfg.relevance_threshold
    vs = get_vectorstore()

    try:
        scored: list[tuple[Document, float]] = vs.similarity_search_with_relevance_scores(
            f"{QUERY_INSTRUCTION}{query}", k=k * 2
        )
    except Exception:
        # 兜底：极少数环境拿不到相关性分数时退回普通检索（不做阈值过滤）
        scored = [(doc, 1.0) for doc in vs.similarity_search(f"{QUERY_INSTRUCTION}{query}", k=k * 2)]

    # 相关性阈值过滤
    candidates = [doc for doc, score in scored if score is None or score >= threshold]

    seen_sources: set[str] = set()
    unique: list[Document] = []
    for doc in candidates:
        source = doc.metadata.get("source", "")
        if source in seen_sources:
            continue
        seen_sources.add(source)
        unique.append(doc)
        if len(unique) >= k:
            break
    return unique
