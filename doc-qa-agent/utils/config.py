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


def _env(name: str, default: str | None = None, required: bool = False) -> str | None:
    """从环境变量取值，支持必填校验。"""
    value = os.getenv(name, default)
    if required and not value:
        raise RuntimeError(f"[config] 缺少必要的环境变量: {name}")
    return value


@dataclass
class Config:
    """全局配置容器。"""

    # ---- LLM ----
    openai_api_key: str = field(default_factory=lambda: _env("OPENAI_API_KEY", required=True))  # type: ignore[assignment]
    openai_base_url: str = field(default_factory=lambda: _env("OPENAI_BASE_URL", required=True))  # type: ignore[assignment]
    model_name: str = field(default_factory=lambda: _env("MODEL_NAME", default="qwen-plus"))  # type: ignore[assignment]

    # ---- Embedding ----
    embedding_model: str = field(default_factory=lambda: _env("EMBEDDING_MODEL", default="BAAI/bge-small-zh-v1.5"))  # type: ignore[assignment]

    # ---- 向量库 ----
    chroma_persist_dir: Path = field(
        default_factory=lambda: Path(_env("CHROMA_PERSIST_DIR", default="./data/chroma")).resolve()
    )

    # ---- 切片 ----
    chunk_size: int = field(default_factory=lambda: int(_env("CHUNK_SIZE", default="512")))  # type: ignore[assignment]
    chunk_overlap: int = field(default_factory=lambda: int(_env("CHUNK_OVERLAP", default="64")))  # type: ignore[assignment]

    # ---- 检索 ----
    retriever_top_k: int = field(default_factory=lambda: int(_env("RETRIEVER_TOP_K", default="4")))  # type: ignore[assignment]

    # ---- 管理口令（上传/清空等管理功能需要；留空则锁定管理功能）----
    admin_password: str = field(default_factory=lambda: _env("ADMIN_PASSWORD", default=""))  # type: ignore[assignment]

    def ensure_dirs(self) -> None:
        """确保所有需要的目录存在。"""
        self.chroma_persist_dir.mkdir(parents=True, exist_ok=True)


# 单例
cfg = Config()
cfg.ensure_dirs()
