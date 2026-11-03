# 企业文档问答助手

一个轻量的**通用企业文档问答知识库 Agent**：上传企业内部文档，基于 LangChain 原生 ReAct Agent 自主检索并作答，答案附带引用来源。

## 适用场景

覆盖任意企业内部文档，例如：

- 管理制度、流程规范
- 产品资料、技术文档
- FAQ、培训材料
- 其他需要集中检索问答的内部文件

不限于行业，也不依赖任何外部联网检索。

## 技术栈

| 组件 | 选型 |
|---|---|
| Web 界面 | Streamlit |
| Agent 框架 | LangChain 原生 `create_react_agent` + `AgentExecutor`（不使用 LangGraph） |
| 向量库 | Chroma（本地持久化） |
| Embedding | sentence-transformers `BAAI/bge-small-zh-v1.5`（本地运行） |
| 文档解析 | pdfplumber（PDF）、python-docx（DOCX）、内置（TXT） |
| 配置 | python-dotenv |
| LLM | 通义千问（OpenAI 兼容接口） |

## 目录结构

```
doc-qa-agent/
├── app.py                      # Streamlit 主入口（侧边栏上传 + 聊天界面）
├── prompts/
│   └── system_prompt.py        # 系统 Prompt（独立维护，便于调整）
├── utils/
│   ├── config.py               # .env 配置加载
│   ├── parsers.py              # PDF / DOCX / TXT 解析
│   ├── vectorstore.py          # Chroma + bge-small-zh 入库 / 检索 / 清空
│   └── agent.py                # ReAct Agent + 唯一工具 knowledge_search
├── requirements.txt
└── .env.example
```

## 快速开始

```powershell
# 1. 创建虚拟环境
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 2. 安装依赖
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

# 3. 配置环境变量
copy .env.example .env
# 编辑 .env，填入 OPENAI_API_KEY / OPENAI_BASE_URL / MODEL_NAME

# 4. 启动
streamlit run app.py
```

首次运行会自动下载 `bge-small-zh-v1.5`（约 90MB）。若网络受限可设置镜像：

```powershell
$env:HF_ENDPOINT = "https://hf-mirror.com"
```

## 功能说明

1. **文档管理**（侧边栏）：上传 PDF / DOCX / TXT、查看已入库文档、清空知识库。
2. **聊天问答**（主区域）：支持多轮对话记忆，Agent 自主决定是否检索、可多次改写关键词检索。
3. **知识库持久化**：向量数据保存在 `./data/chroma`，重启程序不丢失。
4. **引用来源**：回答中标注 `【来源：文档名】`。
5. **不编造**：多次检索仍不足时明确回复「知识库中暂未检索到与该问题相关的资料」，不臆测参数、流程或制度条款。
6. **固定提示**：每次回答末尾附带「本工具仅作辅助参考，请以正式文件与责任人判断为准」。

## 约束与边界

- 仅注册 1 个工具（知识库检索），无联网、无代码执行等其他能力。
- 不做权限管理、OCR、图片识别、Excel 解析，不接 MES。
- 不支持扫描件 PDF（无 OCR）。

## 配置项

| 变量 | 说明 |
|---|---|
| `OPENAI_API_KEY` | LLM API Key（必填） |
| `OPENAI_BASE_URL` | OpenAI 兼容接口地址 |
| `MODEL_NAME` | 模型名，如 `qwen-plus` |
| `EMBEDDING_MODEL` | 本地 Embedding 模型，默认 `BAAI/bge-small-zh-v1.5` |
| `CHROMA_PERSIST_DIR` | 向量库持久化目录，默认 `./data/chroma` |
| `CHUNK_SIZE` / `CHUNK_OVERLAP` | 切片长度与重叠 |
| `RETRIEVER_TOP_K` | 每次检索返回的片段数 |
