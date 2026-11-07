"""
向量库模块
----------
封装 Embedding + Chroma 的初始化、入库、检索、清空逻辑。
向量数据持久化到本地文件夹，重启不丢失。
"""

from __future__ import annotations

import hashlib
import json
import math
import shutil
import threading
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

# 查询同义词扩展：中文制度文件里同一概念常有多种写法（年假 / 年休假、工龄 / 工作年限），
# 短查询直接向量化时会因字面差异排在后面。这里只在命中原词时「追加」同义说法，
# 不改写用户原词，也不引入模型调用；新增领域术语时在这里补充即可。
QUERY_SYNONYMS: dict[str, tuple[str, ...]] = {
    "年假": ("年休假", "年休假天数"),
    "年休假": ("年假", "年休假天数"),
    "工龄": ("工作年限", "本公司工作年限"),
    "工作年限": ("工龄",),
    # 差旅制度里「早中晚」和「早餐/午餐/晚餐」是同一概念的不同说法，
    # 「补助」在制度原文里写作「补贴」，字面不重合会漏召回通用规则片段
    "早中晚": ("早餐", "午餐", "晚餐"),
    "三餐": ("早餐", "午餐", "晚餐"),
    "伙食补助": ("伙食补贴",),
    "餐补": ("伙食补贴", "早餐", "午餐", "晚餐"),
}

# 同一来源最多保留的片段数：
# - 普通文档：3 片（如「年假有几天」这类问题，答案集中在同一文件的相邻段落）
# - Excel：5 片（表格每行语义独立，标准类问题要命中多行才能答全）
MAX_PER_SOURCE = 3
EXCEL_MAX_PER_SOURCE = 5

# Excel 词面补召回：扫描规模上限与最高加分（避免大表拖慢检索）
EXCEL_LEXICAL_SCAN_LIMIT = 3000
EXCEL_LEXICAL_BOOST = 0.25
# 只让「稀有词」参与加权：命中片段占比超过该比例的泛词（如「安全」「住宿」）不加分，
# 否则任何表格行只要带「(含安全)」列就会被抬进结果，产生假阳性
EXCEL_LEXICAL_MAX_DF_RATIO = 0.15
# 单字回退的稀有度更严：单字极易误中（如「理」出现在「交通费处理」），
# 只有真正的专名（香港→「港」、宿迁→「宿」，整张表只出现 1~2 次）才允许加分
EXCEL_LEXICAL_CHAR_MAX_DF = 3
# 「稀有实体精确命中」判定：多字词（如「日本」「北京」）在整张表里只出现不超过这么多次时，
# 视为强证据。表格行是「字段名: 值;」拼接的结构化文本，cross-encoder 对自然语言问法
# 打分极低（实测「日本出差报销多少」→ 日本行 0.06），但实体词本身精确命中足够可靠，
# 这类片段不应被重排阈值直接淘汰。
EXCEL_LEXICAL_STRONG_DF = 3

# 项目根目录（用于把文档路径转换成相对展示位置）
_PROJECT_ROOT = Path(__file__).resolve().parent.parent


def source_path(filename: str) -> str:
    """原始文件的展示位置：相对项目根目录、POSIX 分隔（如 data/docs/xxx.pdf）。"""
    p = cfg.docs_dir / filename
    try:
        return p.relative_to(_PROJECT_ROOT).as_posix()
    except ValueError:
        return p.as_posix()


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


def _make_snippet(content: str, size: int = 20) -> str:
    """
    取出用于在 Word 中定位原文的片段。

    必须保留原文中的空格（Word 的查找是精确匹配，去掉空格会找不到），
    同时不能跨行（跨段落/换行在 Word 里不是连续字符串）。
    """
    for line in content.splitlines():
        text = line.strip()
        if len(text) >= 12:
            return text[:size]
    return content.strip()[:size]


def ingest_document(text: str, filename: str, rows: list[dict] | None = None) -> int:
    """
    将单个文档入库。

    Args:
        text:     文档完整文本（PDF 以 \f 表示分页边界；Excel 为行级切片拼接文本）
        filename: 文档原始文件名（存为 metadata，用于引用）
        rows:     Excel 行级切片（parse_excel_rows 的返回值）。
                  传入时跳过通用字符切片器，直接按「一行一个切片」入库；
                  不传时走原有的 RecursiveCharacterTextSplitter 逻辑。

    Returns:
        成功入库的切片数量

    每个切片除 source 外，还记录定位信息：
        - PDF：page（页码）
        - 其它：line（行号）+ snippet（用于在 Word 中定位的文本片段）
        - Excel：额外的 sheet（工作表名）、row（原始行号）、doc_type="excel"
    """
    docs: list[Document] = []

    if rows is not None:
        # ---- 表格独立路径：不做字符级切分，保留行边界 ----
        for i, row in enumerate(rows):
            content = str(row.get("content") or "").strip()
            if not content:
                continue
            meta: dict = {
                "source": filename,
                "line": int(row.get("line") or (i + 1)),
                "snippet": _make_snippet(content),
                "doc_type": "excel",
            }
            if row.get("sheet"):
                meta["sheet"] = str(row["sheet"])
            if row.get("row"):
                meta["row"] = int(row["row"])
            docs.append(Document(page_content=content, metadata=meta))
    else:
        splitter = _make_splitter()
        chunk_texts = splitter.split_text(text)
        is_pdf = "\f" in text

        cursor = 0
        for chunk in chunk_texts:
            content = chunk.strip()
            if not content:
                continue

            # 定位切片在原文中的起始偏移（切片是原文的连续片段，按顺序向后查找即可）
            start = text.find(chunk, cursor)
            if start < 0:
                start = text.find(chunk)
            if start < 0:
                start = cursor
            cursor = start + len(chunk)

            meta = {"source": filename}
            if is_pdf:
                meta["page"] = text[:start].count("\f") + 1
                content = content.replace("\f", "\n").strip()
            else:
                meta["line"] = text[:start].count("\n") + 1
            meta["snippet"] = _make_snippet(content)

            docs.append(Document(page_content=content, metadata=meta))

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


def _collection_count(vs: Chroma) -> int:
    """集合内切片总数（读不到时返回 0）。"""
    try:
        return int(vs._collection.count())
    except Exception:
        return 0


def _brute_force_search(query: str, k: int) -> list[tuple[Document, float]]:
    """
    兜底检索：不依赖 HNSW 索引，直接对全部切片算余弦相似度。

    用于 HNSW 索引出现 hnswlib 异常（如 "ef or M is too small"）时，
    数据量不大（几千切片内）时开销可接受，可保证问答不中断。
    """
    vs = get_vectorstore()
    data = vs.get(include=["documents", "metadatas", "embeddings"])
    ids = data.get("ids") or []
    docs_raw = data.get("documents") or []
    metas = data.get("metadatas") or []
    vectors = data.get("embeddings")
    if vectors is None or len(ids) == 0:
        return []

    q = get_embedding().embed_query(f"{QUERY_INSTRUCTION}{query}")

    scored: list[tuple[Document, float]] = []
    for i, vec in enumerate(vectors):
        if vec is None:
            continue
        # 入库向量已归一化，余弦相似度等价于点积
        score = float(sum(a * b for a, b in zip(q, vec)))
        doc = Document(page_content=docs_raw[i] or "", metadata=dict(metas[i] or {}))
        scored.append((doc, score))
    scored.sort(key=lambda x: x[1], reverse=True)
    return scored[:k]


def _extract_keywords(query: str, min_kw_len: int = 2) -> list[str]:
    """
    从查询里抽取关键词：中文长串拆 2-gram 滑窗并保留原串，英文单词 / 数字串按原串。
    """
    import re

    chunks = re.findall(r"[\u4e00-\u9fff]+|[A-Za-z]+|\d+", query or "")
    keywords: list[str] = []
    for c in chunks:
        if re.fullmatch(r"[\u4e00-\u9fff]+", c):
            # 同时保留原串（短串就是完整关键词）
            if len(c) >= min_kw_len:
                keywords.append(c)
            if len(c) > min_kw_len:
                for i in range(len(c) - min_kw_len + 1):
                    keywords.append(c[i:i + min_kw_len])
        else:
            keywords.append(c)
    return keywords


def _keyword_overlap_ok(query: str, text: str, min_kw_len: int = 2) -> bool:
    """
    二次校验：向量分数可能被 LLM 扩写的噪声词拉高（比如把"年假"扩成"年假 年休假 工龄"，
    结果和出差补贴/奖惩条例撞分）。这里用简单关键词重叠兜底——query 里连续 ≥min_kw_len
    的中文字串拆成 2-gram 滑窗，只要有一个 gram 在 doc 里出现，就认为有实质关联。
    英文单词 / 数字串按原串匹配。

    注意：短查询（用户原始输入，如「年假有几天」）本身就是精确的，不做词面兜底，
    否则「年假」vs「年休假」这类形态差异会把正确切片误杀。
    """
    if not query or not text:
        return True

    # 短查询跳过过滤：≤16 字的查询基本可以确认是用户原始输入，不含扩写噪声
    if len(query.strip()) <= 16:
        return True

    keywords = _extract_keywords(query, min_kw_len)
    if not keywords:
        return True

    low = text.lower()
    return any(kw.lower() in low for kw in keywords)


# 单字左右两侧的分隔符：只有作为独立字段值出现才算实体命中
_CHAR_LEFT_SEP = ":：、；;，, \t"
_CHAR_RIGHT_SEP = "、；;，,。. \t"


def _is_standalone_char(ch: str, text: str) -> bool:
    """
    判断单字是否作为独立字段值出现，而不是混在普通词语里。

    「国家: 港、澳、台」里的「港」左右都是分隔符 → 是实体简称；
    「明细表」里的「单」左边是「菜」→ 不是实体，不能当强证据。
    """
    start = 0
    while True:
        i = text.find(ch, start)
        if i < 0:
            return False
        left_ok = i == 0 or text[i - 1] in _CHAR_LEFT_SEP
        right_ok = i == len(text) - 1 or text[i + 1] in _CHAR_RIGHT_SEP
        if left_ok and right_ok:
            return True
        start = i + 1


def _excel_lexical_scored(query: str, threshold: float) -> list[tuple[Document, float]]:
    """
    Excel 行级切片的词面补召回（混合检索）。

    bge-small-zh 对「印度」这类实体词区分度很低，表格行又缺少「报销标准」这类泛词，
    实测印度行会掉到向量排名 100 名开外。这里按片段内的词频算 IDF：
    越稀有的命中词权重越高（「印度」只命中 1 条 → 直接顶到最前），
    具体分数抬到阈值之上，再与向量分统一排序。
    """
    keywords = _extract_keywords(query)
    if not keywords:
        return []

    vs = get_vectorstore()
    try:
        data = vs.get(where={"doc_type": "excel"}, include=["documents", "metadatas"])
    except Exception:
        return []

    documents = data.get("documents") or []
    metadatas = data.get("metadatas") or []
    if not documents or len(documents) > EXCEL_LEXICAL_SCAN_LIMIT:
        return []

    total = len(documents)
    lows = [(doc or "").lower() for doc in documents]
    df = {kw: sum(1 for low in lows if kw.lower() in low) for kw in keywords}

    # 单字回退：表格里常用简称（如香港→「港、澳、台」），多字关键词完全无命中时
    # 降级为单字匹配。泛字仍会被下面的稀有度约束拦掉。
    single_chars: list[str] = []
    for kw in keywords:
        if df.get(kw, 0) == 0 and len(kw) >= 2 and all("\u4e00" <= ch <= "\u9fff" for ch in kw):
            single_chars.extend(kw)
    for ch in single_chars:
        if ch not in df:
            df[ch] = sum(1 for low in lows if ch in low)
    single_char_set = set(single_chars)
    keywords = keywords + single_chars

    max_df = max(1, int(total * EXCEL_LEXICAL_MAX_DF_RATIO))
    max_idf = math.log(total + 1)

    scored: list[tuple[Document, float]] = []
    for i, low in enumerate(lows):
        if not low:
            continue
        best_idf = 0.0
        strong = False
        for kw in keywords:
            # 单字走更严的稀有度上限，多字词按比例上限
            kw_max_df = EXCEL_LEXICAL_CHAR_MAX_DF if len(kw) == 1 else max_df
            if df[kw] > kw_max_df:  # 泛词不参与加权
                continue
            if kw.lower() in low:
                best_idf = max(best_idf, math.log((total + 1) / (df[kw] + 1)))
                # 稀有实体精确命中 → 强证据，后续重排阶段即使语义分很低也要保留。
                # 多字词（「日本」「北京」）按 STRONG_DF 判定；
                # 单字只允许「多字词整体没命中后退化出来的简称」（香港→「港」），
                # 且 df≤2 才算专名，避免「理」「程」这类常用字误触发。
                if len(kw) >= 2 and df[kw] <= EXCEL_LEXICAL_STRONG_DF:
                    strong = True
                elif (
                    kw in single_char_set
                    and df[kw] <= 2
                    and _is_standalone_char(kw, low)
                ):
                    strong = True
        if best_idf <= 0:
            continue
        score = threshold + EXCEL_LEXICAL_BOOST * min(1.0, best_idf / max_idf)
        meta = dict(metadatas[i] or {})
        if strong:
            meta["_lexical_strong"] = True
        scored.append((Document(page_content=documents[i] or "", metadata=meta), score))
    return scored


def _expand_query(query: str) -> str:
    """按 QUERY_SYNONYMS 给查询追加同义说法，只扩展不替换原词。"""
    text = (query or "").strip()
    if not text:
        return text
    extras: list[str] = []
    for term, synonyms in QUERY_SYNONYMS.items():
        if term not in text:
            continue
        for synonym in synonyms:
            if synonym not in text and synonym not in extras:
                extras.append(synonym)
    return f"{text} {' '.join(extras)}" if extras else text


# 重排模型：懒加载单例。加载失败（如内网无模型缓存）只记一次标记，
# 后续直接走单阶段检索，不阻塞问答。
_reranker = None
_reranker_failed = False
_rerank_lock = threading.Lock()


def _get_reranker():
    global _reranker, _reranker_failed
    if _reranker_failed:
        return None
    if _reranker is None:
        try:
            from sentence_transformers import CrossEncoder

            _reranker = CrossEncoder(cfg.rerank_model)
        except Exception:
            _reranker_failed = True
            return None
    return _reranker


def _sigmoid(value: float) -> float:
    """把重排模型的 logit 归一化到 0~1，便于设置统一阈值。"""
    if value >= 0:
        return 1.0 / (1.0 + math.exp(-value))
    exp_value = math.exp(value)
    return exp_value / (1.0 + exp_value)


def _rerank_documents(query: str, docs: list[Document]) -> list[Document] | None:
    """
    用 cross-encoder 对候选片段精排。

    返回 None 表示重排不可用（模型缺失 / 推理异常），调用方应回退到单阶段检索；
    返回列表表示重排成功，列表可能为空（全部低于重排阈值）。
    """
    if not docs:
        return []
    model = _get_reranker()
    if model is None:
        return None
    try:
        pairs = [(query, (doc.page_content or "")[:2000]) for doc in docs]
        with _rerank_lock:  # 共享模型实例，避免多会话并发推理
            raw_scores = model.predict(pairs, batch_size=8, show_progress_bar=False)
        # sentence-transformers 的单标签 cross-encoder 默认已内置 Sigmoid，
        # predict 返回的本身就是 0~1 概率（bge-reranker 实测相关 0.93 / 无关 0.0006）。
        # 这里只在模型确实输出越界 logit 时才兜底转换，
        # 否则二次 sigmoid 会把无关片段压到 0.5 附近，让阈值彻底失效。
        scores = [float(score) for score in raw_scores]
        if any(score < 0.0 or score > 1.0 for score in scores):
            scores = [_sigmoid(score) for score in scores]
        ranked = sorted(zip(docs, scores), key=lambda item: item[1], reverse=True)
        return [doc for doc, score in ranked if score >= cfg.rerank_threshold]
    except Exception:
        return None


def _fallback_filter(
    scored: list[tuple[Document, float]],
    threshold: float,
    query: str,
) -> list[Document]:
    """单阶段检索的过滤逻辑：向量阈值 + 关键词重叠，用于重排不可用时降级。"""
    candidates = [doc for doc, score in scored if score is None or score >= threshold]
    return [
        doc
        for doc in candidates
        if doc.metadata.get("doc_type") == "excel"
        or _keyword_overlap_ok(query, doc.page_content or "")
    ]


def similarity_search(query: str, top_k: int | None = None) -> list[Document]:
    """
    根据 query 做两阶段检索，返回 LangChain Document 列表。

    1. 宽召回：向量取 RERANK_CANDIDATES 条候选（不再提前按向量阈值截断），
       并与 Excel IDF 词面补召回结果合并；
    2. 精排：cross-encoder 逐条判断 query 与片段的相关性，sigmoid 归一后按
       RERANK_THRESHOLD 过滤，解决「年假 / 年休假」这类同义表达导致的向量低分。
    重排不可用（模型缺失 / 推理失败）时，自动降级为单阶段检索：
    向量阈值 + 关键词重叠过滤，保证系统始终可用。

    其余规则：
    - 同一来源最多保留 3 个片段（Excel 5 个），保证结果覆盖更多文档；
    - 查询同义词扩展只作为召回辅助，不再承担主要判定职责。
    """
    k = top_k or cfg.retriever_top_k
    threshold = cfg.relevance_threshold
    search_query = _expand_query(query)
    vs = get_vectorstore()

    # 候选数不能超过集合实际切片数，否则 hnswlib 会抛出
    # "Cannot return the results in a contigious 2D array. Probably ef or M is too small"
    # 开启重排时扩大候选池，让向量阶段排名靠后的正确片段也有机会被精排捞回。
    count = _collection_count(vs)
    if count <= 0:
        return []  # 空集合直接返回，避免 hnswlib 在空索引上崩溃
    want = max(k * 4, cfg.rerank_candidates) if cfg.rerank_enabled else k * 4
    n_results = min(want, count)

    scored: list[tuple[Document, float]]
    try:
        scored = vs.similarity_search_with_relevance_scores(
            f"{QUERY_INSTRUCTION}{search_query}", k=n_results
        )
    except Exception:
        try:
            # 兜底 1：普通相似度检索（不带相关性分数）
            scored = [(doc, 1.0) for doc in vs.similarity_search(f"{QUERY_INSTRUCTION}{search_query}", k=n_results)]
        except Exception:
            # 兜底 2：绕过 HNSW 索引，直接全量算相似度
            scored = _brute_force_search(search_query, want)

    # Excel 词面补召回（混合检索）：实体词在向量空间区分度低，靠 IDF 加权把
    # 「印度」这类稀有命中顶到前面；与向量候选合并后统一按分数排序。
    scored.extend(_excel_lexical_scored(query, threshold))  # Excel 补召回用原始 query，避免扩展词误伤
    scored.sort(key=lambda pair: pair[1] if pair[1] is not None else 0.0, reverse=True)

    # 候选按内容去重，减少重排的重复计算（同一表格的相邻切片大量重叠）。
    # 同一内容可能同时来自向量召回和词面补召回，保留先出现（分更高）的那条，
    # 但要把词面补召回标记合并过来，否则去重会丢掉「稀有实体精确命中」的信息。
    deduped: list[Document] = []
    seen_candidates: dict[str, Document] = {}
    for doc, _score in scored:
        content = doc.page_content or ""
        if not content:
            continue
        existed = seen_candidates.get(content)
        if existed is not None:
            if doc.metadata.get("_lexical_strong"):
                existed.metadata["_lexical_strong"] = True
            continue
        seen_candidates[content] = doc
        deduped.append(doc)

    # 两阶段检索：重排器可用时由 cross-encoder 决定相关性；
    # 不可用时降级为「向量阈值 + 关键词重叠」，保证功能不中断。
    reranked = _rerank_documents(query, deduped) if cfg.rerank_enabled else None
    if reranked is None:
        # 降级时把候选池收回到与单阶段一致（k*4），避免宽召回的边缘噪声被阈值放进来
        candidates = _fallback_filter(scored[: k * 4], threshold, search_query)
    else:
        # 稀有实体保底：表格行是「字段名: 值;」拼成的结构化文本，cross-encoder 对
        # 自然语言问法打分极低（实测「日本出差报销多少」→ 日本行 0.06，直接低于阈值被丢）。
        # 但「日本」这类在整张表只出现 1~3 次的实体词精确命中本身就是强证据，
        # 这类片段追加在精排结果之后保留，避免整条正确数据检索不到。
        kept = {id(doc) for doc in reranked}
        strong_extra = [
            doc
            for doc in deduped
            if id(doc) not in kept and doc.metadata.get("_lexical_strong")
        ]
        candidates = reranked + strong_extra[: max(1, k // 2)]

    # Excel 片段需要多条一起才能答全「各职级标准」这类问题：
    # 命中表格时把最终条数上限抬到 EXCEL_MAX_PER_SOURCE（否则会被 top_k=4 压住）
    target = max(k, EXCEL_MAX_PER_SOURCE) if any(
        doc.metadata.get("doc_type") == "excel" for doc in candidates
    ) else k

    # 同一文件保留的片数上限：普通文档 3 片，Excel 5 片；并按内容去重。
    # 两轮策略防止「单一来源高分候选先把 target 占满，其他来源正确片段被截断」：
    #   1) 跨来源保底：候选跨 ≥2 个来源时，每个来源先拿 1 条（取该来源里重排分最高的那条）；
    #      这样「国内各地区」Excel 不会一下子把 5 条额度吃完，境外表（如日本/越南）的正确片段
    #      至少能拿到 1 条保底位置。
    #   2) 再按分数补满：剩余额度继续按重排分顺序填，同时遵守 per-source 上限。
    per_source: dict[str, int] = {}
    seen_contents: set[str] = set()
    unique: list[Document] = []

    # 去重后的候选（按重排分降序，后续两轮都用同一份排序保证确定性）
    dedup: list[Document] = []
    for doc in candidates:
        content = doc.page_content or ""
        if not content or content in seen_contents:
            continue
        seen_contents.add(content)
        dedup.append(doc)

    sources_in_order: list[str] = []
    seen_sources: set[str] = set()
    for doc in dedup:
        src = doc.metadata.get("source", "")
        if src and src not in seen_sources:
            seen_sources.add(src)
            sources_in_order.append(src)

    seen_contents.clear()

    # ---- 1) 跨来源保底 ----
    cross_guard = len(sources_in_order) >= 2
    if cross_guard:
        for src in sources_in_order:
            if len(unique) >= target:
                break
            picked: Document | None = None
            # 优先该来源里「稀有实体精确命中」的片段：同一张境外表可能有「备注」这类
            # 泛化行排在真正的「国家: 日本」行前面，按排序取第一条会保底保错。
            for doc in dedup:
                if doc.metadata.get("source", "") != src:
                    continue
                if not doc.metadata.get("_lexical_strong"):
                    continue
                content = doc.page_content or ""
                if not content or content in seen_contents:
                    continue
                picked = doc
                break
            # 该来源没有实体强命中时，退回排序第一条
            if picked is None:
                for doc in dedup:
                    if doc.metadata.get("source", "") != src:
                        continue
                    content = doc.page_content or ""
                    if not content or content in seen_contents:
                        continue
                    picked = doc
                    break
            if picked is None:
                continue
            seen_contents.add(picked.page_content or "")
            per_source[src] = 1
            unique.append(picked)

    # ---- 2) 剩余额度按重排分补齐 ----
    for doc in dedup:
        if len(unique) >= target:
            break
        content = doc.page_content or ""
        if not content or content in seen_contents:
            continue
        src = doc.metadata.get("source", "")
        limit = (
            EXCEL_MAX_PER_SOURCE
            if doc.metadata.get("doc_type") == "excel"
            else MAX_PER_SOURCE
        )
        if per_source.get(src, 0) >= limit:
            continue
        seen_contents.add(content)
        per_source[src] = per_source.get(src, 0) + 1
        unique.append(doc)

    # 清理内部检索标记，不泄漏给上层（引用展示 / Agent）
    for doc in unique:
        doc.metadata.pop("_lexical_strong", None)

    return unique
