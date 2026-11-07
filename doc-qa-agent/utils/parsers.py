"""
文档解析模块
------------
支持 PDF、DOCX、DOC、TXT、XLSX、XLS 六种文档格式，以及 JPG / PNG 等图片。

三条解析路径：
1. 文本类（PDF / DOCX / DOC / TXT）：parse_file 返回整篇纯文本，交给通用字符切片器；
   PDF 中无文本层的页（扫描件）会自动 OCR，普通电子版 PDF 不受影响；
2. 表格类（XLSX / XLS）：parse_excel_rows 返回「一行一个切片」的独立结构，
   不走通用字符切片器，避免整行语义被字符窗口切断；
3. 图片类（JPG / PNG / BMP / TIFF / WEBP）：整张图片 OCR 后返回纯文本。

对外接口：
    parse_file(file_bytes, filename) -> str
    parse_excel_rows(file_bytes, filename, max_chars=None) -> list[dict]
    rows_to_text(rows) -> str
失败时抛出 DocParseError，由调用方捕获并提示用户。
"""

from __future__ import annotations

import io
import re
import threading
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable

from utils.config import cfg


class DocParseError(Exception):
    """文档解析失败的友好异常。"""


# 支持的扩展名（小写，不含点）
SUPPORTED_EXTENSIONS = {"pdf", "docx", "doc", "txt", "xlsx", "xls", "jpg", "jpeg", "png", "bmp", "tif", "tiff", "webp"}

# 表格类格式：走「行级切片」独立路径
EXCEL_EXTENSIONS = {"xlsx", "xls"}

# 图片格式：整张图 OCR 后入库
IMAGE_EXTENSIONS = {"jpg", "jpeg", "png", "bmp", "tif", "tiff", "webp"}

# 解析进度回调：(已处理页数, 总页数)
ProgressCallback = Callable[[int, int], None]

# 单行拼接后超过该长度时按列分组拆分（默认与 cfg.chunk_size 一致）
_DEFAULT_ROW_MAX_CHARS = 400


# ---------------------------------------------------------------------------
# PDF 解析
# ---------------------------------------------------------------------------
def _parse_pdf(content: bytes, progress: ProgressCallback | None = None) -> str:
    try:
        import pdfplumber
    except ImportError as e:
        raise DocParseError("缺少依赖 pdfplumber，请先执行 `pip install pdfplumber`") from e

    text_parts: list[str] = []
    try:
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            total = len(pdf.pages)
            for index, page in enumerate(pdf.pages, start=1):
                page_text = page.extract_text() or ""

                # 文本层过少视为扫描页：渲染成图片后 OCR，电子版页面完全不受影响
                if len(page_text.strip()) < cfg.ocr_min_text_chars:
                    ocr_text = _ocr_pdf_page(page, index)
                    if ocr_text.strip():
                        page_text = ocr_text

                if page_text.strip():
                    # 用分页符 \f 保留页码边界，供入库时计算切片所在页码
                    text_parts.append(page_text.replace("\f", " ").strip() + "\f")

                if progress is not None:
                    progress(index, total)
    except DocParseError:
        raise
    except Exception as e:
        raise DocParseError(f"PDF 解析失败（文件可能损坏或加密）：{e}") from e

    if not text_parts:
        raise DocParseError(
            "PDF 未提取到任何文本内容（扫描件 OCR 也未识别到文字，"
            "或文件是纯图片且文字不清晰）。"
        )

    return "\n".join(text_parts).rstrip("\f")


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
# DOC 解析（Word 97-2003 二进制格式，需本机安装 Microsoft Word）
# ---------------------------------------------------------------------------
def _parse_doc(content: bytes) -> str:
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as e:
        raise DocParseError("缺少依赖 pywin32，请先执行 `pip install pywin32`") from e

    import tempfile

    with tempfile.TemporaryDirectory() as tmp_dir:
        doc_path = Path(tmp_dir) / "input.doc"
        doc_path.write_bytes(content)

        # COM 要求每个线程显式初始化，否则 Streamlit 子线程调用会失败
        pythoncom.CoInitialize()
        word = None
        try:
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            document = word.Documents.Open(str(doc_path), ReadOnly=True, AddToRecentFiles=False)
            raw_text = document.Content.Text
            document.Close(False)
        except Exception as e:
            raise DocParseError(
                f"DOC 解析失败（本机需安装 Microsoft Word，或另存为 .docx 后上传）：{e}"
            ) from e
        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    # Word 以 \r 换段、\x07 分隔表格单元格、\x0b/\x0c 为软换行/分页
    text = (
        raw_text.replace("\r", "\n")
        .replace("\x07", " | ")
        .replace("\x0b", "\n")
        .replace("\x0c", "\n")
    )
    if not text.strip():
        raise DocParseError("Word 文档未提取到任何文本内容。")
    return text


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
# OCR 解析（扫描件 PDF / 图片）—— 使用本地 RapidOCR，全程离线
# ---------------------------------------------------------------------------
_ocr_engine = None
_ocr_lock = threading.Lock()


def _get_ocr_engine():
    """懒加载 RapidOCR 单例（首次调用约需 1~2 秒加载模型）。"""
    global _ocr_engine
    if _ocr_engine is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
        except ImportError as e:
            raise DocParseError(
                "缺少 OCR 依赖，请先执行 `pip install rapidocr-onnxruntime`"
            ) from e
        try:
            _ocr_engine = RapidOCR()
        except Exception as e:
            raise DocParseError(f"OCR 模型加载失败：{e}") from e
    return _ocr_engine


def _ocr_image(image) -> str:
    """对 PIL 图片做 OCR，按识别顺序返回逐行文本。"""
    import numpy as np

    # RapidOCR 使用 OpenCV 约定，需要 BGR；转成连续内存避免部分版本报错
    rgb = np.asarray(image.convert("RGB"))
    bgr = np.ascontiguousarray(rgb[:, :, ::-1])

    try:
        with _ocr_lock:  # 引擎内部状态非线程安全
            result, _ = _get_ocr_engine()(bgr)
    except DocParseError:
        raise
    except Exception as e:
        raise DocParseError(f"OCR 识别失败：{e}") from e

    if not result:
        return ""

    lines: list[str] = []
    for item in result:
        if len(item) < 2:
            continue
        text = str(item[1]).strip()
        if text:
            lines.append(text)
    return "\n".join(lines)


def _ocr_pdf_page(page, page_no: int) -> str:
    """把 PDF 页渲染为图片后 OCR；渲染失败时给出可读的提示。"""
    try:
        import pypdfium2  # noqa: F401  # pdfplumber 渲染依赖，缺失时提前报错
        image = page.to_image(resolution=cfg.ocr_dpi).original
    except ImportError as e:
        raise DocParseError(
            "缺少 PDF 渲染依赖 pypdfium2，请先执行 `pip install pypdfium2`"
        ) from e
    except Exception as e:
        raise DocParseError(f"第 {page_no} 页扫描图渲染失败：{e}") from e
    return _ocr_image(image)


def _parse_image(content: bytes, progress: ProgressCallback | None = None) -> str:
    """图片文件：整张图 OCR 后返回文本。"""
    try:
        from PIL import Image, ImageOps
    except ImportError as e:
        raise DocParseError("缺少依赖 Pillow，请先执行 `pip install Pillow`") from e

    try:
        with Image.open(io.BytesIO(content)) as image:
            # 手机拍摄的照片常带方向信息，先摆正再识别
            image = ImageOps.exif_transpose(image)
            image.load()
            text = _ocr_image(image)
    except DocParseError:
        raise
    except Exception as e:
        raise DocParseError(f"图片读取失败（文件可能损坏）：{e}") from e

    if progress is not None:
        progress(1, 1)

    if not text.strip():
        raise DocParseError("图片未识别到任何文字（请确认图片清晰、非纯图形）。")
    return text


# ---------------------------------------------------------------------------
# Excel 解析（XLSX / XLS）—— 表格走独立的行级切片路径
# ---------------------------------------------------------------------------
_NUMERIC_RE = re.compile(r"^[¥￥$]?[+-]?\d[\d,]*(?:\.\d+)?[%％]?$")


def _cell_text(value: Any) -> str:
    """把单元格原始值转成适合入库的文本。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "是" if value else "否"
    # 数值：整数型的 float 去掉 .0，避免出现「800.0 元」这种噪声
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, datetime):
        if value.hour or value.minute or value.second:
            return value.strftime("%Y-%m-%d %H:%M")
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    # 单元格内换行/多空格压成单空格，避免破坏行级切片
    return " ".join(str(value).split())


def _looks_numeric(text: str) -> bool:
    return bool(_NUMERIC_RE.match(text.strip()))


def _is_blank_row(cells: list[str]) -> bool:
    return not any(cells)


def _nonempty_count(cells: list[str]) -> int:
    return sum(1 for c in cells if c)


def _is_numeric_row(cells: list[str]) -> bool:
    """数值行判定：非空单元格中过半是纯数值（用于排除「纯数字标题行」）。"""
    values = [c for c in cells if c]
    if len(values) < 2:
        return False
    return sum(1 for c in values if _looks_numeric(c)) * 2 >= len(values)


def _all_same(cells: list[str]) -> bool:
    """整行非空单元格值完全相同 —— 横向合并单元格（标题行）被填充后的特征。"""
    values = [c for c in cells if c]
    return len(values) >= 2 and len(set(values)) == 1


def _has_number_row(cells: list[str]) -> bool:
    """
    数据行判定：非空单元格中至少有一个纯数值。

    不能要求「过半是数值」：很多标准表的列是「城市/职级/金额」，数值只占 1/4，
    按过半判定会把数据行误并进表头，导致列名丢失。
    """
    values = [c for c in cells if c]
    if len(values) < 2:
        return False
    return any(_looks_numeric(c) for c in values)


_YEAR_RE = re.compile(r"^(?:19|20)\d{2}(?:年|年度)?$")


def _header_base(name: str) -> str:
    """去掉重名去重时追加的 _2 / _3 后缀，还原原始列名。"""
    return re.sub(r"_\d+$", "", name) if name else ""


def _is_time_series_sheet(headers: list[str]) -> bool:
    """
    表头过半是年份/纯数字列 → 年度统计或计划表，不是制度/标准类表格，跳过入库。

    例如残留模板表「国家 | 2008年 | 公司愿景 | 2008年」，入库后只会变成噪声。
    自动补名的空列（列3、列4…）不参与占比计算，否则会把统计表的判据稀释掉。
    """
    names = [
        _header_base(h)
        for h in headers
        if h and not re.fullmatch(r"列\d+", h)
    ]
    if len(names) < 2:
        return False
    yearish = sum(1 for h in names if _YEAR_RE.match(h) or _looks_numeric(h))
    return yearish * 2 >= len(names)


def _read_xlsx(content: bytes) -> list[tuple[str, list[list[str]]]]:
    """读取 xlsx，返回 [(工作表名, 二维单元格文本)]。"""
    try:
        from openpyxl import load_workbook
    except ImportError as e:
        raise DocParseError("缺少依赖 openpyxl，请先执行 `pip install openpyxl`") from e

    try:
        book = load_workbook(io.BytesIO(content), data_only=True)
    except Exception as e:
        raise DocParseError(f"Excel（xlsx）解析失败（文件可能损坏或加密）：{e}") from e

    try:
        sheets: list[tuple[str, list[list[str]]]] = []
        for ws in book.worksheets:
            max_row = ws.max_row or 0
            max_col = ws.max_column or 0
            if max_row <= 0 or max_col <= 0:
                continue

            # 合并单元格：openpyxl 只在左上角保留值，必须把该值填充到整个区域，
            # 否则合并区域其余行全为空，语义直接丢失。
            merged: dict[tuple[int, int], str] = {}
            for rng in ws.merged_cells.ranges:
                value = _cell_text(ws.cell(rng.min_row, rng.min_col).value)
                for r in range(rng.min_row, rng.max_row + 1):
                    for c in range(rng.min_col, rng.max_col + 1):
                        merged[(r, c)] = value

            grid: list[list[str]] = []
            for r_idx, row in enumerate(
                ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col, values_only=True),
                start=1,
            ):
                grid.append([
                    merged.get((r_idx, c_idx), _cell_text(value))
                    for c_idx, value in enumerate(row, start=1)
                ])
            sheets.append((ws.title, grid))
        return sheets
    finally:
        try:
            book.close()
        except Exception:
            pass


def _read_xls(content: bytes) -> list[tuple[str, list[list[str]]]]:
    """读取 xls（Word 97-2003 时代的 Excel 格式），返回 [(工作表名, 二维单元格文本)]。"""
    try:
        import xlrd
    except ImportError as e:
        raise DocParseError('缺少依赖 xlrd，请先执行 `pip install "xlrd>=1.2,<2.0"`') from e

    try:
        # formatting_info=True 才会解析合并单元格记录（xlrd 2.x 不支持该参数，降级重试）
        try:
            book = xlrd.open_workbook(file_contents=content, formatting_info=True)
        except TypeError:
            book = xlrd.open_workbook(file_contents=content)
    except Exception as e:
        raise DocParseError(
            f"Excel（xls）解析失败（文件可能损坏），建议用 Excel 另存为 .xlsx 后重新上传：{e}"
        ) from e

    sheets: list[tuple[str, list[list[str]]]] = []
    for sheet in book.sheets():
        grid: list[list[str]] = []
        for r in range(sheet.nrows):
            row: list[str] = []
            for c in range(sheet.ncols):
                cell = sheet.cell(r, c)
                value: Any = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        value = xlrd.xldate_as_datetime(cell.value, book.datemode)
                    except Exception:
                        value = cell.value
                row.append(_cell_text(value))
            grid.append(row)

        # 合并单元格同样需要把左上角值填满整个区域
        for rlo, rhi, clo, chi in getattr(sheet, "merged_cells", []):
            top = grid[rlo][clo] if rlo < len(grid) and clo < len(grid[rlo]) else ""
            for r in range(rlo, min(rhi, len(grid))):
                for c in range(clo, min(chi, len(grid[r]))):
                    grid[r][c] = top

        sheets.append((sheet.name, grid))
    return sheets


def _build_headers(header_rows: list[list[str]], width: int) -> list[str]:
    """把（可能多行的）表头合并成列名：空列名补「列N」，重名加序号。"""
    names: list[str] = []
    for c in range(width):
        parts: list[str] = []
        for row in header_rows:
            value = row[c] if c < len(row) else ""
            if value and (not parts or parts[-1] != value):
                parts.append(value)
        names.append("/".join(parts))

    seen: dict[str, int] = {}
    for i, name in enumerate(names):
        base = name or f"列{i + 1}"
        count = seen.get(base, 0)
        seen[base] = count + 1
        names[i] = base if count == 0 else f"{base}_{count + 1}"
    return names


def _detect_layout(grid: list[list[str]]) -> tuple[list[str], int, str]:
    """
    识别表格结构，返回 (列名列表, 数据起始行索引, 表格标题)。

    规则：
    - 表头之前的纯标题行（整行只有 1 个非空单元格）跳过，第一个标题作为表格名；
    - 以「出现数值行」的位置判定数据区起始，其上方连续非空行作为表头（支持多级表头）；
    - 找不到数值行（全文本表）时退化为「第一行非空行做表头」。
    """
    last = len(grid)
    while last > 0 and _is_blank_row(grid[last - 1]):
        last -= 1
    if last == 0:
        return [], 0, ""

    width = max(len(row) for row in grid)
    max_nonempty = max(_nonempty_count(row) for row in grid)

    idx = 0
    title = ""
    # 单列表格没有「整行仅 1 个非空单元格」的区分意义，不做标题剔除
    if max_nonempty >= 2:
        while idx < last:
            cells = grid[idx]
            nonempty = _nonempty_count(cells)
            if nonempty == 0:
                idx += 1
                continue
            # 标题行两种形态：只有 1 个非空单元格，或横向合并后被填充成整行同值
            if (nonempty == 1 or _all_same(cells)) and not _is_numeric_row(cells):
                title = title or next(c for c in cells if c)
                idx += 1
                continue
            break

    header_start = idx
    if header_start >= last:
        return [], last, title

    data_start: int | None = None
    for j in range(header_start, last):
        if _has_number_row(grid[j]):
            data_start = j
            break

    if data_start is None:
        # 全文本表：第一行当表头
        return _build_headers(grid[header_start:header_start + 1], width), header_start + 1, title

    if data_start == header_start:
        first = grid[header_start][0] if grid[header_start] else ""
        if first and not _looks_numeric(first):
            # 首列是文本、其余列多为数值 → 该行本身就是（年份/数字型多级）表头
            return _build_headers(grid[header_start:header_start + 1], width), data_start + 1, title
        # 首行就是数据行，没有表头
        return [f"列{i + 1}" for i in range(width)], data_start, title

    return _build_headers(grid[header_start:data_start], width), data_start, title


def _row_text(prefix: str, pairs: list[tuple[str, str]]) -> str:
    return prefix + "；".join(f"{key}: {value}" for key, value in pairs)


def _split_pairs(
    prefix: str,
    pairs: list[tuple[str, str]],
    key_pairs: list[tuple[str, str]],
    max_chars: int,
) -> list[str]:
    """单行超长时按列分组拆分；每块都带表格前缀与行主键，保证语义自洽。"""
    full = _row_text(prefix, pairs)
    if len(full) <= max_chars:
        return [full]

    head = _row_text(prefix, key_pairs) + ("；" if key_pairs else "")
    # 行主键已经写进 head，后续分组只处理剩余列，避免同一列重复出现两次
    rest = pairs[len(key_pairs):]
    if not rest:
        return [full]

    pieces: list[str] = []
    buffer: list[tuple[str, str]] = []
    for pair in rest:
        if buffer and len(_row_text(head, buffer + [pair])) > max_chars:
            pieces.append(_row_text(head, buffer))
            buffer = [pair]
        else:
            buffer.append(pair)
    if buffer:
        pieces.append(_row_text(head, buffer))
    return pieces


# 备注类列名关键字：这些列的文本才可能是「全表通用规则」
_NOTE_HEADER_HINTS = ("备注", "说明", "注")

# 通用备注的最短长度与跨行重复门槛：
# 太短的值（「实报实销」「是」「无」）虽然重复很多次，但不是规则条文；
# 只有长度足够且在多行重复出现的文本，才视为被分散填充的通用规则。
_SHARED_NOTE_MIN_CHARS = 10
_SHARED_NOTE_MIN_ROWS = 3


def _collect_shared_notes(
    headers: list[str], grid: list[list[str]], data_start: int
) -> list[tuple[str, int]]:
    """
    收集「在全表多行重复出现」的备注文本。

    部分 Excel 把同一组备注拆成「备注1 / 备注2 / 备注3」多列，每行只填其中一列，
    导致通用规则（如「伙食补贴标准：早餐15元，午餐25元，晚餐30元」）被分散绑定到
    个别地区行上。问其他地区时，模型只能看到「那是青岛的标准」，不敢用于北京。
    这里把重复出现的备注抽出来单列一段，明确标注适用于全表。
    """
    note_cols = [
        i for i, h in enumerate(headers) if any(k in h for k in _NOTE_HEADER_HINTS)
    ]
    if not note_cols:
        return []

    counts: dict[str, int] = {}
    for r in range(data_start, len(grid)):
        cells = grid[r]
        # 按「行」去重：合并单元格展开后同一行的备注 / 备注_2 / 备注_3 值相同，
        # 逐单元格累加会把行数放大数倍，展示给模型的重复行数就不准了。
        row_values = {
            (cells[c] or "").strip()
            for c in note_cols
            if c < len(cells)
        }
        for value in row_values:
            if len(value) >= _SHARED_NOTE_MIN_CHARS:
                counts[value] = counts.get(value, 0) + 1

    # 返回 (文本, 重复行数)：重复次数只作为「可能是通用规则」的线索，
    # 不代表一定适用于全表——也可能是若干地区共用的片区规则。
    return [(text, n) for text, n in counts.items() if n >= _SHARED_NOTE_MIN_ROWS]


def _build_sheet_chunks(sheet_name: str, grid: list[list[str]], max_chars: int) -> list[dict]:
    """把一个工作表转成行级切片列表。"""
    headers, data_start, title = _detect_layout(grid)
    if not headers or data_start >= len(grid):
        return []
    # 年度统计表 / 计划表不是制度类内容，跳过避免污染知识库
    if _is_time_series_sheet(headers):
        return []

    # 前缀带工作表名与表格标题，给每一行提供上下文（避免不同表同名列串味）
    prefix = f"【{sheet_name}】"
    if title:
        prefix += f"{title}｜"

    chunks: list[dict] = []
    width = len(headers)
    for r in range(data_start, len(grid)):
        cells = grid[r]
        if _is_blank_row(cells):
            continue
        pairs: list[tuple[str, str]] = []
        prev_base = ""
        prev_value = ""
        for c in range(min(width, len(cells))):
            value = cells[c]
            if not value:
                continue
            # 合并表头会把同一列拆成「交通工具 / 交通工具_2 / 交通工具_3」，
            # 相邻且列名、列值都相同的列只保留一次，避免同一句话重复三遍。
            base = _header_base(headers[c])
            if base == prev_base and value == prev_value:
                continue
            pairs.append((headers[c], value))
            prev_base, prev_value = base, value
        if not pairs:
            continue
        for content in _split_pairs(prefix, pairs, pairs[:2], max_chars):
            chunks.append({"content": content, "sheet": sheet_name, "row": r + 1})

    # 通用备注单独成片：让「问北京早中晚」这类问题也能命中被分散到其他行的通用规则
    shared_notes = _collect_shared_notes(headers, grid, data_start)
    if shared_notes:
        max_rows = max(n for _text, n in shared_notes)
        chunks.append({
            # 「早中晚三餐」是用户侧常用说法，写进描述里让词面补召回也能命中这段通用规则。
            # 这里只陈述「多行重复出现」这一客观事实，不断言「适用于所有地区」：
            # 重复也可能只是若干地区共用的片区规则，适用范围交给模型结合原文判断。
            "content": prefix
            + f"全表通用说明（早中晚三餐；以下说明在表中最多 {max_rows} 行重复出现，"
            + "请结合具体地区判断适用范围）："
            + "；".join(text for text, _n in shared_notes),
            "sheet": sheet_name,
            "row": 0,
        })
    return chunks


_EXCEL_READERS = {"xlsx": _read_xlsx, "xls": _read_xls}


def parse_excel_rows(
    file_bytes: bytes, filename: str, max_chars: int | None = None
) -> list[dict]:
    """
    解析 Excel，返回行级切片列表（每行即一个独立入库单元）。

    每个元素：{"content": 行文本, "line": 扁平文本行号, "sheet": 工作表名, "row": 原始行号}
    line 与 rows_to_text 生成的行号严格一致，供前端定位与高亮复用。
    """
    ext = Path(filename).suffix.lower().lstrip(".")
    if ext not in EXCEL_EXTENSIONS:
        raise DocParseError(f"{filename} 不是支持的 Excel 格式。")

    sheets = _EXCEL_READERS[ext](file_bytes)
    limit = max_chars if max_chars and max_chars > 0 else _DEFAULT_ROW_MAX_CHARS

    chunks: list[dict] = []
    for sheet_name, grid in sheets:
        chunks.extend(_build_sheet_chunks(sheet_name, grid, limit))

    if not chunks:
        raise DocParseError("Excel 未提取到有效数据行（请确认表格中有内容，且不是纯图片工作表）。")

    for i, item in enumerate(chunks, start=1):
        item["line"] = i
    return chunks


def rows_to_text(rows: list[dict]) -> str:
    """把行级切片拼回扁平文本（与 parse_file 返回值一致，供哈希 / 预览复用）。"""
    return "\n".join(
        str(row.get("content") or "").strip()
        for row in rows
        if str(row.get("content") or "").strip()
    )


# ---------------------------------------------------------------------------
# 统一入口
# ---------------------------------------------------------------------------
_PARSERS = {
    "pdf": _parse_pdf,
    "docx": _parse_docx,
    "doc": _parse_doc,
    "txt": _parse_txt,
}


def parse_file(
    file_bytes: bytes,
    filename: str,
    progress: ProgressCallback | None = None,
) -> str:
    """
    解析单个文件，返回纯文本。

    Args:
        file_bytes: 文件原始字节
        filename:   原始文件名（用于判断扩展名）
        progress:   解析进度回调（目前已处理页数, 总页数），用于扫描件 OCR 的进度展示

    Returns:
        解析后的纯文本字符串

    Raises:
        DocParseError: 不支持的格式 / 文件损坏 / 解析失败
    """
    ext = Path(filename).suffix.lower().lstrip(".")
    if ext not in SUPPORTED_EXTENSIONS:
        raise DocParseError(
            f"暂不支持 {ext.upper()} 格式，仅支持 PDF / DOCX / DOC / TXT / XLSX / XLS "
            "及 JPG / PNG / BMP / TIFF / WEBP 图片。"
        )

    # 表格类：返回行级切片拼接的扁平文本（切片逻辑在 parse_excel_rows 里独立实现）
    if ext in EXCEL_EXTENSIONS:
        return rows_to_text(parse_excel_rows(file_bytes, filename))

    # 图片类：整张图 OCR
    if ext in IMAGE_EXTENSIONS:
        return _parse_image(file_bytes, progress)

    # PDF 需要透传进度回调（扫描页 OCR 可能较慢）
    if ext == "pdf":
        return _parse_pdf(file_bytes, progress)

    return _PARSERS[ext](file_bytes)
