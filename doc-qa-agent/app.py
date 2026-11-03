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

import streamlit as st

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

    st.divider()

    # ---- 清空按钮 ----
    if st.button("🗑️ 清空整个知识库", type="secondary"):
        vectorstore.clear_knowledge_base()
        st.success("知识库已清空。")
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
        with st.spinner("🧠 思考中（可能会检索知识库多次）..."):
            try:
                answer = ask_agent(prompt, history_tuples)
            except Exception as e:
                answer = f"[系统错误] Agent 运行异常：{e}"

        st.markdown(answer)

    st.session_state.messages.append({"role": "assistant", "content": answer})

# 底部提示
st.caption(
    "💡 提示：若知识库尚未上传任何文档，Agent 会如实告知「暂未检索到相关资料」，不会编造内容。"
)
