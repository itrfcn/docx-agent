"""
Streamlit 主程序
----------------
运行: streamlit run app.py

布局：
  左侧侧边栏：文档上传 + 已入库列表 + 清空按钮
  主区域：  多轮聊天界面（会话记忆 + 流式输出）
"""

from __future__ import annotations

import hashlib
import hmac
import os
import subprocess
import sys
import webbrowser
from pathlib import Path

# bge 模型已完整缓存到本地：开启 HF 离线模式，跳过联网检查与重试，
# 避免启动时因无法访问 huggingface.co 反复超时（需在导入 HF 相关库前设置）。
# 若以后要下载新模型：注释掉下一行，即自动走 HF_ENDPOINT 国内镜像下载。
os.environ.setdefault("HF_ENDPOINT", "https://hf-mirror.com")
os.environ.setdefault("HF_HUB_OFFLINE", "1")
os.environ.setdefault("ANONYMIZED_TELEMETRY", "False")

import streamlit as st  # noqa: E402

# 把项目根目录加入 sys.path，保证 `from utils.xxx import ...` 可用
_PROJECT_ROOT = Path(__file__).resolve().parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from utils import parsers, vectorstore  # noqa: E402
from utils.agent import ask_agent, get_last_sources  # noqa: E402
from utils.config import cfg  # noqa: E402


# ---------------------------------------------------------------------------
# 来源文件操作：打开原始文件 / 在资源管理器中定位
# ---------------------------------------------------------------------------
def _resolve_doc_file(name: str) -> Path | None:
    """把来源文件名解析为 docs 目录内的真实文件；越界或不存在返回 None。"""
    base = cfg.docs_dir.resolve()
    path = (base / name).resolve()
    try:
        path.relative_to(base)  # 防止 ../ 之类越出文档目录
    except ValueError:
        return None
    return path if path.is_file() else None


def _open_at_location(path: Path, entry: dict) -> None:
    """打开原始文件并定位到引用位置：PDF 跳到对应页码，其余用 Word 定位到原文片段。"""
    if entry.get("page"):
        webbrowser.open(f"file:///{path.as_posix()}#page={entry['page']}")
        return

    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        os.startfile(str(path))
        return

    pythoncom.CoInitialize()
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = True
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)

        snippet = entry.get("snippet") or ""
        if snippet:
            target = document.Content
            if target.Find.Execute(FindText=snippet, Forward=True, Wrap=1):  # 1 = wdFindContinue
                target.Select()
                word.Activate()
    except Exception as e:
        st.warning(f"已打开文件，但无法定位到引用位置：{e}")
    finally:
        pythoncom.CoUninitialize()


def _location_label(src: dict) -> str:
    """来源定位的简短说明。"""
    if src.get("page"):
        return f"第 {src['page']} 页"
    if src.get("line"):
        return f"第 {src['line']} 行起"
    return ""


def _render_source_actions(sources: list[dict], key_prefix: str) -> None:
    """在回答下方渲染来源文件的操作按钮（仅对磁盘上真实存在的文件显示）。"""
    rows = [(s, _resolve_doc_file(s.get("name", ""))) for s in sources or []]
    rows = [(s, p) for s, p in rows if p is not None]
    if not rows:
        return

    st.caption("原始文件（点击可打开并定位到引用处）")
    for i, (src, path) in enumerate(rows):
        col_open, col_locate = st.columns([0.72, 0.28])
        label = f"📄 {src['name']} · {_location_label(src)}"
        if col_open.button(label, key=f"{key_prefix}_open_{i}", help=str(path)):
            _open_at_location(path, src)
        if col_locate.button("📁 定位", key=f"{key_prefix}_loc_{i}", help="在资源管理器中选中该文件"):
            subprocess.Popen(["explorer", "/select,", str(path)])


# ---------------------------------------------------------------------------
# 页面基础配置
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="企业文档问答助手",
    page_icon="📚",
    layout="wide",
)

# ---- 首屏配置校验 ----
# 缺失必填项时给出友好提示并停止渲染，而不是让未捕获异常导致整页崩溃
_missing_cfg = cfg.missing_required()
if _missing_cfg:
    st.error("❌ 缺少必要配置：" + "、".join(_missing_cfg))
    st.info("请在项目根目录的 .env 文件中填写上述配置后重新启动程序。")
    st.stop()

st.title("📚 企业文档问答助手")
st.caption("面向企业内部员工 · 基于 LangChain ReAct Agent + Chroma + bge-small-zh")

# 隐藏浏览器（Edge/Chromium）给密码框自动加的"显示密码"按钮，
# 避免与 Streamlit 自带的显隐按钮重复出现两个眼睛图标
st.markdown(
    "<style>input::-ms-reveal, input::-ms-clear { display: none; }</style>",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# 侧边栏：文档管理（上传/清空仅管理员可用，文档列表全员可见）
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("📚 文档管理")

    # ---- 管理员认证 ----
    if "is_admin" not in st.session_state:
        st.session_state.is_admin = False

    def _admin_token() -> str:
        """由管理口令派生的访问令牌（口令修改后旧令牌自动失效）。"""
        return hashlib.sha256(("admin-token:" + cfg.admin_password).encode()).hexdigest()[:32]

    # 刷新后恢复登录：URL 中的令牌校验通过即自动恢复管理员状态，
    # 避免每次刷新都要重新输入口令（st.session_state 在刷新后会重建）
    # 注：compare_digest 必须比较 bytes（str 模式不支持非 ASCII 字符）
    if cfg.admin_password and not st.session_state.is_admin:
        qp_token = st.query_params.get("admin_token")
        if qp_token and hmac.compare_digest(
            qp_token.encode("utf-8"), _admin_token().encode("utf-8")
        ):
            st.session_state.is_admin = True

    if not cfg.admin_password:
        st.warning("⚠️ 未在 .env 配置 ADMIN_PASSWORD，管理功能已锁定（仅可查看文档列表）。")
    elif st.session_state.is_admin:
        st.success("🔓 管理员已登录")

    if cfg.admin_password and not st.session_state.is_admin:
        def _attempt_login() -> None:
            """校验输入框中的口令（回车提交与点击按钮共用）。"""
            pwd = st.session_state.get("admin_pwd_input", "")
            if not pwd:
                return
            if hmac.compare_digest(
                pwd.encode("utf-8"), cfg.admin_password.encode("utf-8")
            ):
                st.session_state.is_admin = True
                st.session_state.admin_login_error = False
                # 令牌写入 URL，刷新 / 重开浏览器（同地址）时保持登录
                st.query_params["admin_token"] = _admin_token()
            else:
                st.session_state.admin_login_error = True
            # 清空输入框（回调发生在控件实例化前，允许改其 key）
            st.session_state.admin_pwd_input = ""

        st.text_input(
            "管理员口令",
            type="password",
            key="admin_pwd_input",
            on_change=_attempt_login,  # 输入框内按回车提交即校验
        )
        if st.button("🔓 登录管理", type="primary"):
            _attempt_login()
            if st.session_state.is_admin:
                st.rerun()
        if st.session_state.get("admin_login_error"):
            st.error("口令错误，请重试")

    if cfg.admin_password and st.session_state.is_admin:
        # 清空操作的提示信息（跨 rerun 保留）
        flash_msg = st.session_state.pop("clear_flash", None)
        if flash_msg:
            st.success(flash_msg)

        # ---- 上传组件 ----
        uploaded_files = st.file_uploader(
            "上传文档（可多选）",
            type=list(parsers.SUPPORTED_EXTENSIONS),
            accept_multiple_files=True,
            help="支持 PDF / DOCX / DOC / TXT，不支持扫描件（无 OCR）。.doc 需本机装有 Word。",
        )

        if uploaded_files:
            col_ok, col_clear = st.columns(2)
            if col_ok.button("✅ 入库所选文件", type="primary"):
                progress_bar = st.progress(0.0, text="准备入库...")
                total = len(uploaded_files)
                success_count = 0
                existing = set(vectorstore.list_sources())
                for idx, uf in enumerate(uploaded_files):
                    try:
                        raw = uf.read()
                        text = parsers.parse_file(raw, uf.name)
                        # 原始文件落盘（解析成功后才保存，避免留下损坏文件）；
                        # 仅取文件名，忽略路径成分，防止路径穿越
                        safe_name = Path(uf.name).name
                        cfg.docs_dir.mkdir(parents=True, exist_ok=True)
                        (cfg.docs_dir / safe_name).write_bytes(raw)
                        # 同名覆盖：ingest_document 内部会先删除同名旧切片再写入，
                        # 此处仅用于区分提示文案（解析成功后才写库，失败不影响已有数据）
                        overwritten = uf.name in existing
                        chunks = vectorstore.ingest_document(text, uf.name)
                        success_count += 1
                        if overwritten:
                            st.success(f"🔁 {uf.name} — 已覆盖更新，{chunks} 个切片")
                        else:
                            st.success(f"✅ {uf.name} — 入库 {chunks} 个切片")
                    except parsers.DocParseError as e:
                        st.error(f"❌ {uf.name} — {e}")
                    except Exception as e:
                        st.error(f"❌ {uf.name} — 未知错误：{e}")

                    progress_bar.progress((idx + 1) / total, text=f"已处理 {idx + 1}/{total}")

                progress_bar.empty()
                st.info(f"入库完成：{success_count}/{total} 个文件成功。")

            if col_clear.button("↩️ 取消选择"):
                st.rerun()

        # ---- 清空按钮（二次确认，防误删）----
        st.divider()
        if "confirm_clear" not in st.session_state:
            st.session_state.confirm_clear = False

        if not st.session_state.confirm_clear:
            # 第一次点击只进入确认态，不执行清空
            if st.button("🗑️ 清空整个知识库", type="secondary"):
                st.session_state.confirm_clear = True
                st.rerun()
        else:
            st.warning("⚠️ 此操作会删除全部已入库文档，且无法恢复。")
            confirmed = st.checkbox("我确认要清空全部已入库文档")

            col_yes, col_no = st.columns(2)
            with col_yes:
                do_clear = st.button(
                    "确认清空",
                    type="primary",
                    disabled=not confirmed,
                )
            with col_no:
                if st.button("取消"):
                    st.session_state.confirm_clear = False
                    st.rerun()

            if do_clear:
                # 二次确认通过后才真正执行
                result = vectorstore.clear_knowledge_base()
                docs = result.get("documents") or []
                st.session_state.confirm_clear = False  # 执行完自动复位

                if result.get("cleared"):
                    msg = (
                        "✅ 知识库已清空：向量集合已重建，"
                        f"manifest 文档清单共 {len(docs)} 个"
                    )
                    msg += ("：" + "、".join(docs)) if docs else "。"
                    if result.get("leftover"):
                        msg += (
                            "\n\nℹ️ 向量数据已清空，残留文件将在关闭程序后自动清理，不影响使用。"
                        )
                    st.session_state.clear_flash = msg
                else:
                    st.session_state.clear_flash = (
                        "❌ 清空失败：知识库可能正被其它程序占用，请稍后重试。"
                    )
                st.rerun()

        # ---- 退出管理 ----
        if st.button("🔒 退出管理"):
            st.session_state.is_admin = False
            # 同时清除 URL 中的令牌，避免刷新后又被自动登录
            if "admin_token" in st.query_params:
                del st.query_params["admin_token"]
            st.rerun()

        # 版本/配置提示（仅管理员可见）
        st.divider()
        with st.expander("🔧 当前配置"):
            st.write(f"LLM: `{cfg.model_name}`")
            st.write(f"Embedding: `{cfg.embedding_model}`")
            st.write(f"向量库: `{cfg.chroma_persist_dir}`")
            st.write(f"切片: size={cfg.chunk_size}, overlap={cfg.chunk_overlap}")

    st.divider()

    # ---- 已入库文件列表（全员可见；删除仅管理员）----
    st.subheader("📂 已入库文档")
    try:
        sources = vectorstore.list_sources()
    except Exception as e:
        st.error(f"读取知识库失败：{e}")
        sources = []

    is_admin_view = bool(cfg.admin_password) and st.session_state.get("is_admin", False)
    if "pending_delete" not in st.session_state:
        st.session_state.pending_delete = None

    if sources:
        for src in sources:
            if not is_admin_view:
                st.markdown(f"- `{src}`")
                continue

            col_name, col_del = st.columns([0.82, 0.18])
            col_name.markdown(f"- `{src}`")
            if col_del.button("🗑", key=f"del_{src}", help="删除该文档的全部切片"):
                st.session_state.pending_delete = src

            # 二次确认（仅对触发删除的那个文件显示）
            if st.session_state.pending_delete == src:
                st.warning(f"删除 `{src}` 的全部切片？此操作不可恢复。")
                c_yes, c_no = st.columns(2)
                if c_yes.button("确认删除", key=f"delok_{src}", type="primary"):
                    try:
                        n = vectorstore.delete_document(src)
                        st.session_state.pending_delete = None
                        st.session_state.clear_flash = f"✅ 已删除 {src}（{n} 个切片），其余文档不受影响"
                    except Exception as e:
                        st.session_state.pending_delete = None
                        st.session_state.clear_flash = f"❌ 删除失败：{e}"
                    st.rerun()
                if c_no.button("取消", key=f"delno_{src}"):
                    st.session_state.pending_delete = None
                    st.rerun()
        st.caption(f"共 {len(sources)} 个文档")
    else:
        st.caption("（知识库为空，请先上传文档）")


# ---------------------------------------------------------------------------
# 主区域：聊天界面
# ---------------------------------------------------------------------------
# 初始化会话状态
if "messages" not in st.session_state:
    st.session_state.messages = []  # list of {"role": ..., "content": ...}

# 渲染历史消息
for idx, msg in enumerate(st.session_state.messages):
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])
        if msg.get("sources"):
            _render_source_actions(msg["sources"], f"hist{idx}")

# 聊天输入框
prompt = st.chat_input("请输入问题（如：差旅报销的审批流程是什么？）")
if prompt:
    # 显示用户消息
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # 组装历史（human, ai 交替元组）
    history_tuples: list[tuple[str, str]] = []
    buf_human: str | None = None
    for m in st.session_state.messages[:-1]:  # 排除刚加的 user
        if m["role"] == "user":
            buf_human = m["content"]
        elif m["role"] == "assistant" and buf_human is not None:
            history_tuples.append((buf_human, m["content"]))
            buf_human = None

    # Agent 回复
    with st.chat_message("assistant"):
        with st.spinner("🧠 思考中（知识库检索中,请耐心等待）..."):
            try:
                answer = ask_agent(prompt, history_tuples)
            except Exception:
                # 不把技术细节写进对话历史，只给用户友好提示
                answer = "抱歉，当前无法完成回答，请稍后重试。"
            hit_sources = get_last_sources()

        st.markdown(answer)
        _render_source_actions(hit_sources, "live")

    st.session_state.messages.append(
        {"role": "assistant", "content": answer, "sources": hit_sources}
    )

# 底部提示
st.caption(
    "💡 提示：Agent 只能回答知识库内的资料，若知识库未检索到相关资料，将不会回答。"
)
