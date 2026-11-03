"""
Streamlit 主程序
----------------
运行: streamlit run app.py

布局：
  左侧侧边栏：文档上传 + 已入库列表 + 清空按钮
  主区域：  多轮聊天界面（会话记忆 + 流式输出）
"""

from __future__ import annotations

import sys
from pathlib import Path

import streamlit as st  # noqa: E402

# 把项目根目录加入 sys.path，保证 `from utils.xxx import ...` 可用
_PROJECT_ROOT = Path(__file__).resolve().parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from utils import parsers, vectorstore  # noqa: E402
from utils.agent import ask_agent  # noqa: E402
from utils.config import cfg  # noqa: E402


# ---------------------------------------------------------------------------
# 页面基础配置
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="企业文档问答助手",
    page_icon="📚",
    layout="wide",
)

st.title("📚 企业文档问答助手")
st.caption("面向企业内部员工 · 基于 LangChain ReAct Agent + Chroma + bge-small-zh")

# ---------------------------------------------------------------------------
# 侧边栏：文档管理
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("📚 文档管理")

    # 清空操作的提示信息（跨 rerun 保留）
    flash_msg = st.session_state.pop("clear_flash", None)
    if flash_msg:
        st.success(flash_msg)

    # ---- 上传组件 ----
    uploaded_files = st.file_uploader(
        "上传文档（可多选）",
        type=list(parsers.SUPPORTED_EXTENSIONS),
        accept_multiple_files=True,
        help="支持 PDF / DOCX / TXT，不支持扫描件（无 OCR）。",
    )

    if uploaded_files:
        col_ok, col_clear = st.columns(2)
        if col_ok.button("✅ 入库所选文件", type="primary"):
            progress_bar = st.progress(0.0, text="准备入库...")
            total = len(uploaded_files)
            success_count = 0
            for idx, uf in enumerate(uploaded_files):
                try:
                    raw = uf.read()
                    text = parsers.parse_file(raw, uf.name)
                    chunks = vectorstore.ingest_document(text, uf.name)
                    success_count += 1
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

    st.divider()

    # ---- 已入库文件列表 ----
    st.subheader("📂 已入库文档")
    try:
        sources = vectorstore.list_sources()
    except Exception as e:
        st.error(f"读取知识库失败：{e}")
        sources = []

    if sources:
        for src in sources:
            st.markdown(f"- `{src}`")
        st.caption(f"共 {len(sources)} 个文档")
    else:
        st.caption("（知识库为空，请先上传文档）")

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

    # 版本/配置提示
    st.divider()
    with st.expander("🔧 当前配置"):
        st.write(f"LLM: `{cfg.model_name}`")
        st.write(f"Embedding: `{cfg.embedding_model}`")
        st.write(f"向量库: `{cfg.chroma_persist_dir}`")
        st.write(f"切片: size={cfg.chunk_size}, overlap={cfg.chunk_overlap}")


# ---------------------------------------------------------------------------
# 主区域：聊天界面
# ---------------------------------------------------------------------------
# 初始化会话状态
if "messages" not in st.session_state:
    st.session_state.messages = []  # list of {"role": ..., "content": ...}

# 渲染历史消息
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

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
            except Exception as e:
                answer = f"[系统错误] Agent 运行异常：{e}"

        st.markdown(answer)

    st.session_state.messages.append({"role": "assistant", "content": answer})

# 底部提示
st.caption(
    "💡 提示：Agent 只能回答知识库内的资料，若知识库未检索到相关资料，将不会回答。"
)
