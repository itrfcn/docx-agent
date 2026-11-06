"""
浏览器预览模块
--------------
把原始文档转成 PDF / HTML 缓存到 static/preview，供浏览器直接打开。

三大优化（v3）：
1. 加密标记持久化 — 写 JSON 文件，Streamlit 重启后也知道哪些文件被 DLP 加密
2. Word 进程常驻 — 单例 + 锁，避免每次 DispatchEx 冷启动（~2.5s → ~0.15s）
3. 后台预热 — 用户还在看 Agent 回答时，后台线程悄悄把 read_text 跑完

静态目录需配合 .streamlit/config.toml 的 server.enableStaticServing = true。
"""

from __future__ import annotations

import hashlib
import html
import json
import os
import shutil
import threading
import time
import uuid
import atexit
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
STATIC_PREVIEW_DIR = _PROJECT_ROOT / "static" / "preview"

# ---------------------------------------------------------------------------
# 清理策略
# ---------------------------------------------------------------------------
_TTL_SECONDS = 24 * 60 * 60
_MAX_FILES = 200
_MAX_BYTES = 50 * 1024 * 1024
_last_cleanup_ts = 0.0

# ---------------------------------------------------------------------------
# 缓存
# ---------------------------------------------------------------------------
# 进程内版本级缓存（key = 文件名 + mtime + size 前 12 位 SHA1）
_VERSION_TEXT_CACHE: dict[str, str] = {}

# 被企业加密的文件版本（内存 + 磁盘双存）
_ENCRYPTED_FILE = STATIC_PREVIEW_DIR / ".encrypted.json"
_VERSION_PDF_ENCRYPTED: dict[str, bool] = {}

# ---------------------------------------------------------------------------
# Word 单例常驻：避免每次 DispatchEx 冷启动
# ---------------------------------------------------------------------------
_word_app = None
_word_lock = threading.Lock()  # 所有 Word COM 操作共用一把锁（STA 模型）

# Word COM 常量
_WD_FORMAT_PDF = 17
_WD_ACTIVE_END_PAGE_NUMBER = 3
_WD_READ_ONLY = True
_WD_ADD_TO_RECENT = False


def _encrypted_load() -> dict[str, bool]:
    try:
        if _ENCRYPTED_FILE.exists():
            data = json.loads(_ENCRYPTED_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return {k: bool(v) for k, v in data.items()}
    except OSError:
        pass
    return {}


def _encrypted_save() -> None:
    try:
        STATIC_PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
        _ENCRYPTED_FILE.write_text(
            json.dumps(_VERSION_PDF_ENCRYPTED, ensure_ascii=False),
            encoding="utf-8",
        )
    except OSError:
        pass


# 模块导入时先把磁盘上的加密标记加载进来
_VERSION_PDF_ENCRYPTED = _encrypted_load()


def _mark_encrypted(key: str) -> None:
    """标记该版本已被 DLP 加密，下次直接跳过 PDF 导出。"""
    _VERSION_PDF_ENCRYPTED[key] = True
    _encrypted_save()


# ---------------------------------------------------------------------------
# Word 进程常驻管理
# ---------------------------------------------------------------------------
def _word_ensure_up():
    """确保 Word 进程已启动；调用方已持有 _word_lock 且已 CoInitialize。"""
    global _word_app
    if _word_app is None:
        try:
            import win32com.client as win32
        except ImportError:
            return None
        _word_app = win32.DispatchEx("Word.Application")
        _word_app.Visible = False
        _word_app.DisplayAlerts = 0
    return _word_app


def _word_quit_idle() -> None:
    """退出 Word 释放内存（被 atexit 在进程退出时调用）。"""
    global _word_app
    if _word_app is None:
        return
    try:
        _word_app.Quit()
    except Exception:
        pass
    _word_app = None


@atexit.register
def _word_quit_on_exit() -> None:
    """Python 进程退出时清理常驻 Word。"""
    with _word_lock:
        _word_quit_idle()


def _run_word_operation(fn):
    """
    包装一次 Word COM 操作：CoInitialize → 持锁 → fn(app) → 解锁 → CoUninitialize。
    发生 Word 级异常时自动重试一次（Word 可能崩溃了）。
    """
    import pythoncom
    global _word_app
    last_error = None
    for attempt in range(2):
        pythoncom.CoInitialize()
        try:
            with _word_lock:
                app = _word_ensure_up()
                if app is None:
                    return None
                result = fn(app)
            return result
        except Exception as e:
            last_error = e
            # Word 可能崩溃了，强制重置
            try:
                with _word_lock:
                    _word_app = None
            except Exception:
                _word_app = None
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        if attempt == 0:
            continue
    raise last_error  # type: ignore[misc]


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------
def _touch(path: Path) -> None:
    try:
        now = time.time()
        os.utime(path, (now, now))
    except OSError:
        pass


def _cleanup_cache() -> None:
    """机会式清理（同进程 10 分钟最多扫一次）。"""
    global _last_cleanup_ts
    now = time.time()
    if now - _last_cleanup_ts < 600:
        return
    _last_cleanup_ts = now

    try:
        files = [p for p in STATIC_PREVIEW_DIR.iterdir() if p.is_file()]
    except OSError:
        return

    if not files:
        return

    files.sort(key=lambda p: p.stat().st_atime)

    # TTL 过期
    for p in list(files):
        if now - p.stat().st_atime > _TTL_SECONDS:
            try:
                p.unlink()
            except OSError:
                pass
    files = [p for p in files if p.exists()]

    # 总量上限
    total_bytes = sum(p.stat().st_size for p in files)
    while len(files) > _MAX_FILES or total_bytes > _MAX_BYTES:
        if not files:
            break
        victim = files.pop(0)
        total_bytes -= victim.stat().st_size
        try:
            victim.unlink()
        except OSError:
            pass


def _cache_key(path: Path) -> str:
    stat = path.stat()
    return hashlib.sha1(
        f"{path.name}:{stat.st_mtime_ns}:{stat.st_size}".encode()
    ).hexdigest()[:12]


def _is_valid_pdf(path: Path) -> bool:
    try:
        with path.open("rb") as f:
            return f.read(5) == b"%PDF-"
    except OSError:
        return False


# ---------------------------------------------------------------------------
# Word → PDF 导出（使用常驻 Word）
# ---------------------------------------------------------------------------
def _export_pdf(src: Path, out: Path, snippet: str) -> int | None:
    """用常驻 Word 导出 PDF 并返回片段所在页码；失败返回 None，不产生半成品。"""
    src = src.resolve()
    tmp = out.with_name(f"{out.stem}.{uuid.uuid4().hex}.tmp.pdf")

    def _do_word(app):
        if app is None:
            return None
        document = app.Documents.Open(
            str(src), ReadOnly=_WD_READ_ONLY, AddToRecentFiles=_WD_ADD_TO_RECENT
        )
        page: int | None = None
        if snippet:
            target = document.Content
            if target.Find.Execute(FindText=snippet, Forward=True, Wrap=1):
                page = int(target.Information(_WD_ACTIVE_END_PAGE_NUMBER))
        document.SaveAs2(str(tmp), FileFormat=_WD_FORMAT_PDF)
        document.Close(False)
        return page

    try:
        page = _run_word_operation(_do_word)
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


# ---------------------------------------------------------------------------
# 对外 API
# ---------------------------------------------------------------------------
def pdf_preview_url(path: Path, entry: dict) -> str | None:
    """生成浏览器可直接打开的 PDF 预览地址（带页码锚点）。"""
    try:
        from utils import parsers

        # Excel / 图片没有 Word/PDF 版式预览：直接返回 None 走网页全文预览。
        # 必须在这里提前返回，否则 Word 导出失败会被误标成「企业加密」文件。
        ext = path.suffix.lower().lstrip(".")
        if ext in parsers.EXCEL_EXTENSIONS or ext in parsers.IMAGE_EXTENSIONS:
            return None

        STATIC_PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
        _cleanup_cache()
        key = _cache_key(path)
        out = STATIC_PREVIEW_DIR / f"{key}.pdf"

        page: int | None = int(entry["page"]) if entry.get("page") else None

        if path.suffix.lower() != ".pdf" and _VERSION_PDF_ENCRYPTED.get(key):
            return None

        if path.suffix.lower() == ".pdf":
            if not out.exists():
                shutil.copy2(path, out)
        else:
            found = _export_pdf(path, out, str(entry.get("snippet") or ""))
            if found is not None:
                page = found

        if not out.exists() or not _is_valid_pdf(out):
            _mark_encrypted(key)
            try:
                if out.exists():
                    out.unlink()
            except OSError:
                pass
            return None

        _touch(out)
        url = f"/app/static/preview/{out.name}"
        return f"{url}#page={page}" if page else url
    except Exception:
        return None


def read_text(path: Path) -> str:
    """
    读取文档纯文本：优先用解析器；失败时改用常驻 Word（可读取被企业加密的文档）。
    同进程内按文件版本缓存，命中即返回毫秒级。
    """
    cache_key = _cache_key(path)
    if cache_key in _VERSION_TEXT_CACHE:
        return _VERSION_TEXT_CACHE[cache_key]

    from utils import parsers

    text = ""
    try:
        text = parsers.parse_file(path.read_bytes(), path.name)
        if text.strip():
            _VERSION_TEXT_CACHE[cache_key] = text
            return text
    except Exception:
        pass

    # Excel / 图片无法用 Word 兜底读取，直接返回空（避免无意义的 Word COM 调用）
    ext = path.suffix.lower().lstrip(".")
    if ext in parsers.EXCEL_EXTENSIONS or ext in parsers.IMAGE_EXTENSIONS:
        _VERSION_TEXT_CACHE[cache_key] = ""
        return ""

    # 走 Word 常驻：Open → Content.Text → Close
    def _do_word(app):
        if app is None:
            return ""
        doc = app.Documents.Open(
            str(path.resolve()),
            ReadOnly=_WD_READ_ONLY,
            AddToRecentFiles=_WD_ADD_TO_RECENT,
        )
        try:
            return doc.Content.Text
        finally:
            try:
                doc.Close(False)
            except Exception:
                pass

    try:
        raw = _run_word_operation(_do_word) or ""
    except Exception:
        raw = ""

    text = (
        raw.replace("\r", "\n")
        .replace("\x07", " | ")
        .replace("\x0b", "\n")
        .replace("\x0c", "\n")
    )
    _VERSION_TEXT_CACHE[cache_key] = text
    return text


def html_preview_url(path: Path, entry: dict) -> str | None:
    """生成全文网页预览地址（引用片段高亮并自动滚动到该处）。"""
    try:
        text = read_text(path)
        if not text.strip():
            return None

        STATIC_PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
        _cleanup_cache()
        page_no = entry.get("page")
        line_no = entry.get("line")
        loc = int(page_no or line_no or 0)
        out = STATIC_PREVIEW_DIR / f"{_cache_key(path)}-{loc}.html"

        if not out.exists():
            body = html.escape(text)
            snippet = html.escape(str(entry.get("snippet") or ""))
            if snippet:
                # 优先高亮「行号对应的那一行」：Excel / 文本类切片的片段前缀往往重复，
                # 直接替换第一次出现会永远定位到第 1 行
                lines = body.split("\n")
                target = int(line_no or 0) - 1
                if 0 <= target < len(lines) and snippet in lines[target]:
                    lines[target] = lines[target].replace(
                        snippet, f'<mark id="hit">{snippet}</mark>', 1
                    )
                    body = "\n".join(lines)
                elif snippet in body:
                    body = body.replace(snippet, f'<mark id="hit">{snippet}</mark>', 1)

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

        _touch(out)
        return f"/app/static/preview/{out.name}"
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 后台预热
# ---------------------------------------------------------------------------
def warmup_sources(names: list[str], docs_dir: Path) -> None:
    """
    后台多线程预热：Agent 返回回答后立刻跑，让用户点预览按钮时 read_text / PDF 导出
    直接命中缓存。

    - 用 set 去重，避免同一文件多次预热
    - 每个文件一个 daemon 线程，不阻塞主线程、不阻塞进程退出
    - 线程内部吞掉异常，不影响主线程
    """
    from utils import parsers

    unique_paths = [
        docs_dir.resolve() / n
        for n in {str(n).strip() for n in names if n}
        if (docs_dir / n).is_file()
    ]

    def _one(p: Path) -> None:
        try:
            read_text(p)
        except Exception:
            pass
        # 顺带把 PDF 也导一遍（首次 SaveAs2 ~0.5s，之后完全命中缓存）
        try:
            ext = p.suffix.lower().lstrip(".")
            if (
                ext != "pdf"
                and ext not in parsers.EXCEL_EXTENSIONS
                and ext not in parsers.IMAGE_EXTENSIONS
            ):
                key = _cache_key(p)
                out = STATIC_PREVIEW_DIR / f"{key}.pdf"
                if not out.exists() and not _VERSION_PDF_ENCRYPTED.get(key):
                    _export_pdf(p, out, "")
        except Exception:
            pass

    for p in unique_paths:
        threading.Thread(target=_one, args=(p,), daemon=True).start()
