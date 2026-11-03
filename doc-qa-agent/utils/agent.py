"""
Agent 模块
----------
使用 LangChain 原生 ReAct Agent（create_react_agent）。
仅注册一个工具：知识库检索工具 knowledge_search。
"""

from __future__ import annotations

from typing import Any

from langchain.agents import create_react_agent, AgentExecutor
from langchain_core.tools import tool
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_core.messages import SystemMessage

from utils.config import cfg
from utils import vectorstore
from prompts.system_prompt import SYSTEM_PROMPT


# ---------------------------------------------------------------------------
# 1. 知识库检索工具（唯一工具）
# ---------------------------------------------------------------------------
@tool
def knowledge_search(query: str) -> str:
    """在企业内部文档知识库中检索与问题相关的文档片段。
    当用户询问企业管理制度、流程规范、产品与技术资料、FAQ、培训材料等业务问题时必须调用此工具。
    输入：清晰、具体的检索关键词或问题。
    输出：检索到的文档片段（含来源文件名），若无相关内容会明确说明。"""
    try:
        docs = vectorstore.similarity_search(query)
    except Exception as e:
        return f"[知识库检索异常] {e}"

    if not docs:
        return "知识库中未检索到相关片段（可能没有相关资料）。"

    # 去重（按内容）并格式化，附带来源
    seen: set[str] = set()
    parts: list[str] = []
    for doc in docs:
        content = (doc.page_content or "").strip()
        if not content or content in seen:
            continue
        seen.add(content)
        source = doc.metadata.get("source", "未知来源")
        parts.append(f"[{source}]\n{content}")

    if not parts:
        return "知识库中未检索到相关片段（可能没有相关资料）。"

    return "\n\n---\n\n".join(parts)


# ---------------------------------------------------------------------------
# 2. 构造 Agent（延迟初始化，避免 Streamlit 热重载时重复加载模型）
# ---------------------------------------------------------------------------
_agent_executor: AgentExecutor | None = None
_llm: ChatOpenAI | None = None


def _get_llm() -> ChatOpenAI:
    """LLM 单例：60 秒超时 + 2 次重试，避免单次请求长时间卡住。"""
    global _llm
    if _llm is None:
        _llm = ChatOpenAI(
            api_key=cfg.openai_api_key,
            base_url=cfg.openai_base_url,
            model=cfg.model_name,
            temperature=0.1,
            streaming=False,
            timeout=60,
            max_retries=2,
        )
    return _llm


def get_agent_executor() -> AgentExecutor:
    """获取 ReAct Agent 单例。"""
    global _agent_executor
    if _agent_executor is not None:
        return _agent_executor

    # 关闭 streaming，确保模型一次返回完整输出，避免 ReAct 格式被截断
    llm = _get_llm()

    tools = [knowledge_search]

    # 使用 ReAct 标准提示模板（含历史消息占位）
    # 注意：create_react_agent 要求：
    #   - 必须包含 {tools} / {tool_names} 变量
    #   - agent_scratchpad 是字符串（内部用 format_log_to_str 生成），不能用 MessagesPlaceholder
    system_content = (
        SYSTEM_PROMPT
        + "\n\n## 可用工具\n"
        + "工具名称: {tool_names}\n"
        + "工具详情:\n{tools}"
    )
    prompt = ChatPromptTemplate.from_messages([
        ("system", system_content),
        MessagesPlaceholder(variable_name="chat_history"),
        ("human", "{input}\n\nThought:{agent_scratchpad}"),
    ])

    agent = create_react_agent(llm=llm, tools=tools, prompt=prompt)

    _agent_executor = AgentExecutor(
        agent=agent,
        tools=tools,
        verbose=True,              # 控制台打印思考过程（生产可关）
        handle_parsing_errors=(
            "输出格式错误：同一条回复不能同时包含 Action 和 Final Answer。"
            "请重新输出——要么只输出 Thought/Action/Action Input 三行后停止，要么只输出 Final Answer。"
        ),  # 解析失败时把精确原因反馈给模型，避免盲目重试
        max_iterations=10,         # 防止死循环，允许 10 轮检索内收敛
        max_execution_time=120,    # 整体执行超时（秒），超时后进入降级直答
        # 注：当前 langchain 版本仅支持 early_stopping_method="force"（默认），不可用 "generate"
        return_intermediate_steps=False,
    )
    return _agent_executor


# ---------------------------------------------------------------------------
# 3. 对外调用接口
# ---------------------------------------------------------------------------
_DISCLAIMER = "\n\n---\n⚠️ **本工具仅作辅助参考，请以正式文件与责任人判断为准。**"


def _is_degraded_answer(text: str) -> bool:
    """识别 ReAct 未正常收敛的情况（迭代/时间耗尽、空输出）。"""
    t = (text or "").strip().lower()
    return (not t) or ("iteration limit" in t) or ("time limit" in t) or ("agent stopped" in t)


def _direct_answer(question: str) -> str:
    """
    降级路径：直接检索知识库 → 一次 LLM 直答。

    ReAct 解析失败 / 迭代耗尽 / 请求超时时使用，避免把英文技术报错暴露给用户。
    """
    try:
        docs = vectorstore.similarity_search(question)
    except Exception:
        docs = []

    if not docs:
        return "抱歉，知识库中暂未检索到与该问题相关的资料，无法回答该问题。" + _DISCLAIMER

    context_parts: list[str] = []
    source_names: list[str] = []
    for doc in docs:
        source = doc.metadata.get("source", "未知来源")
        content = (doc.page_content or "").strip()
        if not content:
            continue
        context_parts.append(f"[{source}]\n{content}")
        if source not in source_names:
            source_names.append(source)

    if not context_parts:
        return "抱歉，知识库中暂未检索到与该问题相关的资料，无法回答该问题。" + _DISCLAIMER

    prompt = ChatPromptTemplate.from_messages([
        (
            "system",
            "你是企业内部文档问答助手。只能依据用户提供的资料回答；"
            "资料不足时必须明确说明未检索到相关资料，严禁编造制度、流程或数据。"
            "用简洁中文分点陈述，每个结论后用【来源：文件名】标注来源。",
        ),
        ("human", "问题：{question}\n\n资料：\n{context}"),
    ])

    try:
        msg = _get_llm().invoke(
            prompt.format_messages(question=question, context="\n\n".join(context_parts))
        )
        answer = str(getattr(msg, "content", "") or "").strip()
        if not answer:
            raise ValueError("empty answer")
    except Exception:
        return "抱歉，当前无法完成回答，请稍后重试。" + _DISCLAIMER

    if "【来源" not in answer and "【引用来源" not in answer:
        answer += f"\n\n【引用来源：{'、'.join(source_names)}】"
    if "本工具仅作辅助参考" not in answer:
        answer += _DISCLAIMER
    return answer


def ask_agent(user_question: str, chat_history: list[tuple[str, str]] | None = None) -> str:
    """
    向 Agent 发起一次问答。

    Args:
        user_question: 用户当前提问
        chat_history:  历史对话，格式 [(human_msg, ai_msg), ...]

    Returns:
        Agent 最终回复文本（任何异常都会降级为友好中文答复，不返回技术报错）
    """
    executor = get_agent_executor()

    # LangChain 的 chat_history 使用 BaseMessage 序列
    from langchain_core.messages import HumanMessage, AIMessage

    lc_history: list = []
    for human_msg, ai_msg in (chat_history or []):
        lc_history.append(HumanMessage(content=human_msg))
        lc_history.append(AIMessage(content=ai_msg))

    try:
        result = executor.invoke({
            "input": user_question,
            "chat_history": lc_history,
        })
        answer = str(result.get("output", "")).strip()
    except Exception:
        # 超时 / 限流 / 解析失败等：不把英文技术报错暴露给用户
        answer = ""

    if _is_degraded_answer(answer):
        answer = _direct_answer(user_question)
    return answer
