"""
Agent 模块
----------
使用 LangChain 原生 ReAct Agent（create_react_agent）。
仅注册一个工具：知识库检索工具 knowledge_search。
"""

from __future__ import annotations

import threading
from typing import Any

from langchain.agents import create_react_agent, AgentExecutor
from langchain.agents.output_parsers import ReActSingleInputOutputParser
from langchain_core.agents import AgentFinish
from langchain_core.exceptions import OutputParserException
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
# 本轮问答命中的来源文件（线程本地：Streamlit 每个会话独立线程，互不干扰）
_source_ctx = threading.local()


def _record_source(doc: Any) -> None:
    """记录本轮检索命中的文档位置（同一文件的多个不同位置都会保留）。"""
    meta: dict = dict(getattr(doc, "metadata", {}) or {})
    source = meta.get("source", "未知来源")

    if meta.get("page"):
        loc = {"page": int(meta["page"])}
    else:
        loc = {
            "line": int(meta.get("line") or 0),
            "snippet": str(meta.get("snippet") or ""),
        }
    entry: dict = {"name": source, **loc}

    items: list[dict] = getattr(_source_ctx, "items", []) or []
    # 同一文件、同一位置只记一次（按位置去重，而非按文件名去重）
    for item in items:
        if item.get("name") == source and {
            k: v for k, v in item.items() if k != "name"
        } == loc:
            return

    items.append(entry)
    _source_ctx.items = items


def get_last_sources() -> list[dict]:
    """返回最近一次 ask_agent 命中的来源文件列表：{"name": ..., "page"|"line": ...}。"""
    return list(getattr(_source_ctx, "items", []) or [])


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

    # 去重（按内容）并格式化，附带来源文件名
    seen: set[str] = set()
    parts: list[str] = []
    for doc in docs:
        content = (doc.page_content or "").strip()
        if not content or content in seen:
            continue
        seen.add(content)
        source = doc.metadata.get("source", "未知来源")
        _record_source(doc)
        parts.append(f"[{source}]\n{content}")

    if not parts:
        return "知识库中未检索到相关片段（可能没有相关资料）。"

    return "\n\n---\n\n".join(parts)


# ---------------------------------------------------------------------------
# 2. 构造 Agent（延迟初始化，避免 Streamlit 热重载时重复加载模型）
# ---------------------------------------------------------------------------
_agent_executor: AgentExecutor | None = None
_llm: ChatOpenAI | None = None


class _LenientReActParser(ReActSingleInputOutputParser):
    """
    宽松版 ReAct 解析器。

    背景：部分模型习惯直接给出答案正文而漏写 "Final Answer:" 前缀，
    标准解析器会抛 OutputParserException，导致白白多跑一轮 LLM 调用。

    策略：
    - 正常格式 → 交由标准解析器处理（行为完全不变）；
    - 解析失败，但文本里**没有** Action 意图、且内容足够长 → 直接视为最终答案；
    - 其余情况（想调工具却格式混乱、输出过短）→ 继续抛异常走纠正重试。
    """

    _MIN_ANSWER_LEN = 20
    # 以这些字符结尾说明输出被截断/未写完，不能当作完整答案
    _INCOMPLETE_ENDINGS = ("，", ",", "、", "：", ":", "（", "(", "【", "…")

    def parse(self, text: str):
        try:
            return super().parse(text)
        except OutputParserException:
            body = (text or "").strip()
            # 出现 Action 说明模型仍想调用工具，只是格式乱了，需要纠正后重试
            if "Action:" in body or "Action Input:" in body:
                raise
            if len(body) < self._MIN_ANSWER_LEN:
                raise
            if body.endswith(self._INCOMPLETE_ENDINGS):
                raise
            # 去掉可能残留的 "Thought:" 前缀，保留正文
            if body.startswith("Thought:"):
                body = body[len("Thought:"):].lstrip()
            return AgentFinish(return_values={"output": body}, log=text)


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

    agent = create_react_agent(
        llm=llm,
        tools=tools,
        prompt=prompt,
        output_parser=_LenientReActParser(),   # 漏写 Final Answer 前缀时直接采纳正文
    )

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
        # 需要中间步骤来判断「是否真正调用过检索工具」（硬性规则见 _needs_force_retrieval）
        return_intermediate_steps=True,
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


# 明确声明“资料不足”的说法：这类回答不需要来源标注
_REFUSAL_HINTS = ("暂无相关", "未检索到", "没有相关", "无法回答")


def _looks_grounded(text: str) -> bool:
    """
    判断答案是否基于知识库。

    宽松解析器会直接采纳缺少 "Final Answer:" 前缀的回答，存在一种风险：
    模型没检索就凭常识作答。这里用两点兜底：
    - 带来源标注（【来源 / 【引用来源）→ 视为有依据；
    - 明确声明资料不足 → 视为合规回答。
    都没有且篇幅较长（≥40 字，排除闲聊短句）→ 判为无依据，转强制检索直答。
    """
    if "【来源" in text or "【引用来源" in text:
        return True
    if any(hint in text for hint in _REFUSAL_HINTS):
        return True
    return len(text.strip()) < 40


def _needs_force_retrieval(answer: str, steps: list) -> bool:
    """
    硬性规则：实质性回答必须建立在一次真实检索之上。

    - 迭代/时间耗尽、空输出 → 需兜底；
    - 无来源标注、非明确拒答的长答案 → 需兜底；
    - 有来源标注但整轮**从未调用过检索工具** → 来源可能是编造的，需兜底。
    命中后由 _direct_answer 执行「检索 + 基于资料作答」。
    """
    if _is_degraded_answer(answer):
        return True
    if not _looks_grounded(answer):
        return True
    if not steps and len(answer.strip()) >= 40:
        return True
    return False


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
    source_labels: list[str] = []
    for doc in docs:
        source = doc.metadata.get("source", "未知来源")
        content = (doc.page_content or "").strip()
        if not content:
            continue
        _record_source(doc)
        context_parts.append(f"[{source}]\n{content}")
        if source not in source_labels:
            source_labels.append(source)

    if not context_parts:
        return "抱歉，知识库中暂未检索到与该问题相关的资料，无法回答该问题。" + _DISCLAIMER

    prompt = ChatPromptTemplate.from_messages([
        (
            "system",
            "你是企业内部文档问答助手。只能依据用户提供的资料回答；"
            "资料不足时必须明确说明未检索到相关资料，严禁编造制度、流程或数据。"
            "用简洁中文分点陈述，每个结论后用【来源：文件名】标注来源，只写文件名，不要附加文件路径。",
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
        answer += f"\n\n【引用来源：{'、'.join(source_labels)}】"
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
    _source_ctx.items = []  # 重置本轮来源记录

    # LangChain 的 chat_history 使用 BaseMessage 序列
    from langchain_core.messages import HumanMessage, AIMessage

    lc_history: list = []
    for human_msg, ai_msg in (chat_history or []):
        lc_history.append(HumanMessage(content=human_msg))
        lc_history.append(AIMessage(content=ai_msg))

    steps: list = []
    try:
        result = executor.invoke({
            "input": user_question,
            "chat_history": lc_history,
        })
        answer = str(result.get("output", "")).strip()
        steps = result.get("intermediate_steps") or []
    except Exception:
        # 超时 / 限流 / 解析失败等：不把英文技术报错暴露给用户
        answer = ""

    if _needs_force_retrieval(answer, steps):
        answer = _direct_answer(user_question)
    return answer
