"""
本地文件夹批量导入
------------------
用途：企业环境无法使用网页上传时，把 PDF/DOCX/TXT 放进 data/docs 目录，
      执行本脚本即可批量入库（复用 解析 → 切片 → embedding → Chroma 写入 链路）。

用法：
    python import_docs.py

规则：
    - 递归扫描 data/docs 下的 pdf / docx / txt
    - 已入库文件自动跳过（以 Chroma 中已有的 source 名作为去重依据）
    - 单个文件失败不影响其余文件
    - 不删除、不覆盖原有已入库文档
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

# 国内网络可走 HuggingFace 镜像（已存在该环境变量时不覆盖）
os.environ.setdefault("HF_ENDPOINT", "https://hf-mirror.com")
os.environ.setdefault("ANONYMIZED_TELEMETRY", "False")

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from utils import parsers, vectorstore  # noqa: E402

# 待扫描目录（可通过环境变量 IMPORT_DOCS_DIR 覆盖）
DOCS_DIR = Path(os.getenv("IMPORT_DOCS_DIR", PROJECT_ROOT / "data" / "docs")).resolve()


def _source_name(path: Path) -> str:
    """入库时使用的文档名：相对 data/docs 的路径（POSIX 风格，兼顾子目录）。"""
    try:
        return path.relative_to(DOCS_DIR).as_posix()
    except ValueError:
        return path.name


def _iter_doc_files(root: Path) -> list[Path]:
    """递归收集受支持的文档文件（跳过 Word 临时锁文件 ~$xxx）。"""
    files: list[Path] = []
    for p in sorted(root.rglob("*")):
        if p.name.startswith("~$"):
            continue
        if p.is_file() and p.suffix.lower().lstrip(".") in parsers.SUPPORTED_EXTENSIONS:
            files.append(p)
    return files


def main() -> int:
    print("=" * 64)
    print("本地文档批量导入")
    print(f"扫描目录：{DOCS_DIR}")
    print("=" * 64)

    if not DOCS_DIR.exists():
        print(f"⚠️  目录不存在：{DOCS_DIR}")
        print("   请先创建该目录，并把 pdf / docx / txt 文件放进去。")
        return 1

    files = _iter_doc_files(DOCS_DIR)
    if not files:
        print("⚠️  目录为空（未找到 pdf / docx / txt 文件）。")
        print("   请把文档放入后重新执行：python import_docs.py")
        return 0

    # 已入库文档名（沿用现有 Chroma metadata 作为 manifest）
    try:
        existing = set(vectorstore.list_sources())
    except Exception as e:
        print(f"❌ 读取现有知识库失败：{e}")
        print("   请确认 data/chroma 未被其它程序占用后重试。")
        return 1

    print(f"发现 {len(files)} 个文件，知识库中已有 {len(existing)} 个文档\n")

    added, skipped, failed = 0, 0, 0
    details: list[tuple[str, str, str]] = []  # (状态, 文件名, 说明)

    for path in files:
        name = _source_name(path)
        try:
            if name in existing:
                skipped += 1
                details.append(("跳过", name, "已入库"))
                print(f"[跳过] {name} —— 已入库")
                continue

            raw = path.read_bytes()
            text = parsers.parse_file(raw, path.name)
            chunks = vectorstore.ingest_document(text, name)
            if chunks <= 0:
                raise RuntimeError("解析后无有效文本切片")

            added += 1
            existing.add(name)
            details.append(("新增", name, f"{chunks} 个切片"))
            print(f"[新增] {name} —— {chunks} 个切片")
        except parsers.DocParseError as e:
            failed += 1
            details.append(("失败", name, str(e)))
            print(f"[失败] {name} —— {e}")
        except Exception as e:  # 单文件失败不影响其余文件
            failed += 1
            details.append(("失败", name, f"未知错误：{e}"))
            print(f"[失败] {name} —— 未知错误：{e}")

    print("\n" + "-" * 64)
    print(f"汇总：新增 {added} 个 / 跳过 {skipped} 个 / 失败 {failed} 个")
    if failed:
        print("失败明细：")
        for status, name, reason in details:
            if status == "失败":
                print(f"  - {name}：{reason}")

    print("\n提示：若 Streamlit 正在运行，刷新页面即可看到新文档；")
    print("      若侧边栏仍不显示，重启 streamlit 服务即可（数据已持久化，不会丢失）。")
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
