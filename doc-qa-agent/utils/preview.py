"""
浏览器预览模块
--------------
把原始文档转成 PDF 并缓存到 static/preview，供浏览器直接打开。

- PDF：直接复制到静态目录（浏览器原生渲染）
- DOC/DOCX：借助本机 Word（COM）导出 PDF，并用 Word 查找原文片段得到真实页码

静态目录需配合 .streamlit/config.toml 中的 server.enableStaticServing = true，
浏览器访问路径为 /app/static/preview/<缓存文件>.pdf。
"""

from __future__ import annotations

import hashlib
import html
import shutil
import uuid
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
STATIC_PREVIEW_DIR = _PROJECT_ROOT / "static" / "preview"

# Word 导出格式：wdFormatPDF
_WD_FORMAT_PDF = 17
# Range.Information 参数：wdActiveEndPageNumber
_WD_ACTIVE_END_PAGE_NUMBER = 3


def _export_pdf(src: Path, out: Path, snippet: str) -> int | None:
    """
    用 Word 把文档导出为 PDF，并返回片段所在的真实页码。

    转换失败（未安装 Word / COM 异常）返回 None，且不产生半成品文件。
    """
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        return None

    pythoncom.CoInitialize()
    word = None
    src = src.resolve()  # Word 无法解析相对路径，必须传绝对路径
    tmp = out.with_name(f"{out.stem}.{uuid.uuid4().hex}.tmp.pdf")
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)

        page: int | None = None
        if snippet:
            target = document.Content
            if target.Find.Execute(FindText=snippet, Forward=True, Wrap=1):  # 1 = wdFindContinue
                page = int(target.Information(_WD_ACTIVE_END_PAGE_NUMBER))

        document.SaveAs2(str(tmp), FileFormat=_WD_FORMAT_PDF)
        document.Close(False)
        if not tmp.exists():
            return None
        tmp.replace(out)
        return page
    except Exception:
        return None
    finally:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _cache_key(path: Path) -> str:
    """按文件版本生成缓存键（文件名 + 修改时间 + 大小）。"""
    stat = path.stat()
    return hashlib.sha1(
        f"{path.name}:{stat.st_mtime_ns}:{stat.st_size}".encode()
    ).hexdigest()[:12]


def _is_valid_pdf(path: Path) -> bool:
    """检查文件是否真的是 PDF。

    企业加密软件（DLP）会把 Word 写出的文件加密，Python/浏览器读到的是密文，
    这类文件没有 %PDF 文件头，不能交给浏览器预览。
    """
    try:
        with path.open("rb") as f:
            return f.read(5) == b"%PDF-"
    except OSError:
        return False


def pdf_preview_url(path: Path, entry: dict) -> str | None:
    """
    生成浏览器可直接打开的 PDF 预览地址（带页码锚点）。

    Args:
        path:  原始文件路径（data/docs 下）
        entry: 来源定位信息，可能含 page / snippet

    Returns:
        形如 /app/static/preview/xxx.pdf#page=3 的地址；无法转换或文件被加密时返回 None
    """
    try:
        STATIC_PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
        key = _cache_key(path)
        out = STATIC_PREVIEW_DIR / f"{key}.pdf"

        page: int | None = int(entry["page"]) if entry.get("page") else None

        if path.suffix.lower() == ".pdf":
            if not out.exists():
                shutil.copy2(path, out)
        else:
            found = _export_pdf(path, out, str(entry.get("snippet") or ""))
            if found:
                page = found

        if not out.exists() or not _is_valid_pdf(out):
            # Word 导出的文件被企业加密时不是真正的 PDF，清理后交给 HTML 预览兜底
            try:
                if out.exists():
                    out.unlink()
            except OSError:
                pass
            return None

        url = f"/app/static/preview/{out.name}"
        return f"{url}#page={page}" if page else url
    except Exception:
        return None


def read_text(path: Path) -> str:
    """
    读取文档纯文本：优先用解析器；失败时改用 Word（可读取被企业加密的 DOC/DOCX）。
    """
    from utils import parsers

    try:
        text = parsers.parse_file(path.read_bytes(), path.name)
        if text.strip():
            return text
    except Exception:
        pass

    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        return ""

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(
            str(path.resolve()), ReadOnly=True, AddToRecentFiles=False
        )
        raw = document.Content.Text
        document.Close(False)
    except Exception:
        return ""
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return (
        raw.replace("\r", "\n")
        .replace("\x07", " | ")
        .replace("\x0b", "\n")
        .replace("\x0c", "\n")
    )


def html_preview_url(path: Path, entry: dict) -> str | None:
    """
    生成全文网页预览地址（引用片段高亮并自动滚动到该处）。

    由 Python 直接写出 HTML，避免 Word 导出文件被企业加密后无法被浏览器读取。
    """
    try:
        text = read_text(path)
        if not text.strip():
            return None

        STATIC_PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
        page_no = entry.get("page")
        line_no = entry.get("line")
        loc = int(page_no or line_no or 0)
        out = STATIC_PREVIEW_DIR / f"{_cache_key(path)}-{loc}.html"

        if not out.exists():
            body = html.escape(text)
            snippet = html.escape(str(entry.get("snippet") or ""))
            if snippet and snippet in body:
                body = body.replace(
                    snippet, f'<mark id="hit">{snippet}</mark>', 1
                )

            if page_no:
                loc_text = f"第 {int(page_no)} 页"
            elif line_no:
                loc_text = f"第 {int(line_no)} 行起"
            else:
                loc_text = "全文"

            title = html.escape(path.name)
            out.write_text(
                f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>{title}</title>
<style>
body {{ margin: 0; background: #eef0f3; color: #1f2328;
        font-family: "Microsoft YaHei", "PingFang SC", sans-serif; }}
.bar {{ position: sticky; top: 0; z-index: 1; padding: 12px 24px;
        background: #fff; border-bottom: 1px solid #e3e6ea;
        font-size: 14px; color: #555; }}
.doc {{ max-width: 900px; margin: 20px auto; padding: 32px 40px;
        background: #fff; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,.08);
        white-space: pre-wrap; line-height: 1.9; font-size: 15px; }}
mark {{ background: #ffe58f; padding: 0 2px; }}
</style>
</head>
<body>
<div class="bar">{title} · 定位：{loc_text}（黄色为引用片段）</div>
<div class="doc">{body}</div>
<script>
const hit = document.getElementById('hit');
if (hit) hit.scrollIntoView({{ block: 'center' }});
</script>
</body>
</html>""",
                encoding="utf-8",
            )

        return f"/app/static/preview/{out.name}"
    except Exception:
        return None
