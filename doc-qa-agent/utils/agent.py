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


def get_agent_executor() -> AgentExecutor:
    """获取 ReAct Agent 单例。"""
    global _agent_executor
    if _agent_executor is not None:
        return _agent_executor

    # LLM：通义千问兼容 OpenAI 接口
    # 关闭 streaming，确保模型一次返回完整输出，避免 ReAct 格式被截断
    llm = ChatOpenAI(
        api_key=cfg.openai_api_key,
        base_url=cfg.openai_base_url,
        model=cfg.model_name,
        temperature=0.1,
        streaming=False,
    )

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
        # 注：当前 langchain 版本仅支持 early_stopping_method="force"（默认），不可用 "generate"
        return_intermediate_steps=False,
    )
    return _agent_executor


# ---------------------------------------------------------------------------
# 3. 对外调用接口
# ---------------------------------------------------------------------------
def ask_agent(user_question: str, chat_history: list[tuple[str, str]] | None = None) -> str:
    """
    向 Agent 发起一次问答。

    Args:
        user_question: 用户当前提问
        chat_history:  历史对话，格式 [(human_msg, ai_msg), ...]

    Returns:
        Agent 最终回复文本
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
    except Exception as e:
        return f"[Agent 执行异常] 抱歉，推理过程中出现错误：{e}"

    answer = result.get("output", "")
    return str(answer).strip()
