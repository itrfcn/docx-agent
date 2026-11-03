"""
文档解析模块
------------
支持 PDF、DOCX、TXT 三种格式。
对外暴露统一接口：parse_file(file_bytes: bytes, filename: str) -> str
失败时抛出 DocParseError，由调用方捕获并提示用户。
"""

from __future__ import annotations

import io
from pathlib import Path


class DocParseError(Exception):
    """文档解析失败的友好异常。"""


# 支持的扩展名（小写，不含点）
SUPPORTED_EXTENSIONS = {"pdf", "docx", "txt"}


# ---------------------------------------------------------------------------
# PDF 解析
# ---------------------------------------------------------------------------
def _parse_pdf(content: bytes) -> str:
    try:
        import pdfplumber
    except ImportError as e:
        raise DocParseError("缺少依赖 pdfplumber，请先执行 `pip install pdfplumber`") from e

    text_parts: list[str] = []
    try:
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                if page_text.strip():
                    text_parts.append(page_text)
    except Exception as e:
        raise DocParseError(f"PDF 解析失败（文件可能损坏或加密）：{e}") from e

    if not text_parts:
        raise DocParseError("PDF 未提取到任何文本内容（可能是纯扫描件，本工具不支持 OCR）。")

    return "\n".join(text_parts)


# ---------------------------------------------------------------------------
# DOCX 解析
# ---------------------------------------------------------------------------
def _parse_docx(content: bytes) -> str:
    try:
        from docx import Document
    except ImportError as e:
        raise DocParseError("缺少依赖 python-docx，请先执行 `pip install python-docx`") from e

    try:
        doc = Document(io.BytesIO(content))
    except Exception as e:
        raise DocParseError(f"Word 文档解析失败（文件可能损坏）：{e}") from e

    text_parts: list[str] = [para.text for para in doc.paragraphs if para.text.strip()]
    # 顺带提取表格文本（可选，提升召回率）
    for table in doc.tables:
        for row in table.rows:
            row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
            if row_text:
                text_parts.append(row_text)

    if not text_parts:
        raise DocParseError("Word 文档未提取到任何文本内容。")

    return "\n".join(text_parts)


# ---------------------------------------------------------------------------
# TXT 解析
# ---------------------------------------------------------------------------
def _parse_txt(content: bytes) -> str:
    # 尝试 UTF-8，回退到 GBK（中文 Windows 常见编码）
    for encoding in ("utf-8", "gbk"):
        try:
            return content.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise DocParseError("TXT 文件编码无法识别（已尝试 UTF-8 和 GBK）。")


# ---------------------------------------------------------------------------
# 统一入口
# ---------------------------------------------------------------------------
_PARSERS = {
    "pdf": _parse_pdf,
    "docx": _parse_docx,
    "txt": _parse_txt,
}


def parse_file(file_bytes: bytes, filename: str) -> str:
    """
    解析单个文件，返回纯文本。

    Args:
        file_bytes: 文件原始字节
        filename:   原始文件名（用于判断扩展名）

    Returns:
        解析后的纯文本字符串

    Raises:
        DocParseError: 不支持的格式 / 文件损坏 / 解析失败
    """
    ext = Path(filename).suffix.lower().lstrip(".")
    if ext not in SUPPORTED_EXTENSIONS:
        raise DocParseError(f"暂不支持 {ext.upper()} 格式，仅支持 PDF / DOCX / TXT。")

    parser = _PARSERS[ext]
    return parser(file_bytes)
