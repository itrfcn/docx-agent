"""
向量库模块
----------
封装 Embedding + Chroma 的初始化、入库、检索、清空逻辑。
向量数据持久化到本地文件夹，重启不丢失。
"""

from __future__ import annotations

from pathlib import Path

from langchain_core.documents import Document
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_chroma import Chroma
from langchain_huggingface import HuggingFaceEmbeddings

from utils.config import cfg

# Chroma collection 名称（业务隔离）
COLLECTION_NAME = "enterprise_docs"

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


def clear_knowledge_base() -> None:
    """
    清空整个知识库（删除当前 collection）。
    直接删除 persist_dir 下的数据，保证彻底。
    """
    global _vectorstore
    if _vectorstore is not None:
        _vectorstore.delete_collection()
        _vectorstore = None

    # 同时尝试从磁盘物理清空
    persist_dir: Path = cfg.chroma_persist_dir
    if persist_dir.exists():
        # 递归删除目录下所有文件/子目录
        for child in persist_dir.iterdir():
            if child.is_file() or child.is_symlink():
                child.unlink(missing_ok=True)
            elif child.is_dir():
                import shutil
                shutil.rmtree(child, ignore_errors=True)


def similarity_search(query: str, top_k: int | None = None) -> list[Document]:
    """
    根据 query 做相似度检索，返回 LangChain Document 列表。
    """
    k = top_k or cfg.retriever_top_k
    vs = get_vectorstore()
    return vs.similarity_search(query, k=k)
