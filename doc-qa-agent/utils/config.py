"""
配置加载模块
------------
从 .env 文件读取环境变量，并暴露统一的 Config 类供全局使用。
所有路径类配置均在应用启动时解析为绝对路径。
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from dotenv import load_dotenv

# 加载 .env（优先读取项目根目录下的 .env）
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(_PROJECT_ROOT / ".env")


def _env(name: str, default: str | None = None) -> str | None:
    """从环境变量取值。

    注意：这里不做必填校验，避免配置缺失时在 import 阶段直接抛异常导致整页崩溃；
    必填项由 Config.missing_required() 收集，页面首屏用友好提示展示。
    """
    return os.getenv(name, default)


@dataclass
class Config:
    """全局配置容器。"""

    # ---- LLM ----
    # 必填项不做导入期校验，缺失时由 missing_required() 统一提示
    openai_api_key: str = field(default_factory=lambda: _env("OPENAI_API_KEY", default="") or "")
    openai_base_url: str = field(default_factory=lambda: _env("OPENAI_BASE_URL", default="") or "")
    model_name: str = field(default_factory=lambda: _env("MODEL_NAME", default="qwen-plus") or "qwen-plus")

    # ---- Embedding ----
    embedding_model: str = field(default_factory=lambda: _env("EMBEDDING_MODEL", default="BAAI/bge-small-zh-v1.5") or "BAAI/bge-small-zh-v1.5")

    # ---- 向量库 ----
    chroma_persist_dir: Path = field(
        default_factory=lambda: Path(_env("CHROMA_PERSIST_DIR", default="./data/chroma")).resolve()
    )

    # ---- 切片 ----
    # 400 字左右更适合中文制度/流程文档，命中片段更聚焦；可通过环境变量调整
    chunk_size: int = field(default_factory=lambda: int(_env("CHUNK_SIZE", default="400")))  # type: ignore[assignment]
    chunk_overlap: int = field(default_factory=lambda: int(_env("CHUNK_OVERLAP", default="64")))  # type: ignore[assignment]

    # ---- 检索 ----
    retriever_top_k: int = field(default_factory=lambda: int(_env("RETRIEVER_TOP_K", default="4")))  # type: ignore[assignment]
    # 相关性阈值（0~1，越高越严格）：低于该分数的片段不返回。
    # cosine 空间下 relevance = 1 - 余弦距离，0.45 为经验起点。
    relevance_threshold: float = field(
        default_factory=lambda: float(_env("RELEVANCE_THRESHOLD", default="0.45"))
    )

    # ---- 管理口令（上传/清空等管理功能需要；留空则锁定管理功能）----
    admin_password: str = field(default_factory=lambda: _env("ADMIN_PASSWORD", default="") or "")

    def missing_required(self) -> list[str]:
        """返回缺失的必填配置项名称列表（供页面首屏友好提示）。"""
        missing: list[str] = []
        if not self.openai_api_key:
            missing.append("OPENAI_API_KEY")
        if not self.openai_base_url:
            missing.append("OPENAI_BASE_URL")
        return missing

    def ensure_dirs(self) -> None:
        """确保所有需要的目录存在。"""
        self.chroma_persist_dir.mkdir(parents=True, exist_ok=True)


# 单例
cfg = Config()
cfg.ensure_dirs()
