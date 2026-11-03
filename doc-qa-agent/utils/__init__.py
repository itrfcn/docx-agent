"""
utils 包初始化
--------------
在此统一设置 HuggingFace 相关环境变量（必须先于任何 langchain / transformers
模块导入执行：`from utils.xxx import ...` 一定会先运行本文件）。

- bge 模型已完整缓存到本地：开启 HF 离线模式，跳过联网检查与重试，
  避免网络不通时 tokenizer/processor 加载异常导致检索工具报错。
  （huggingface_hub 在被导入那一刻读取 HF_HUB_OFFLINE，之后再设无效，
  而 agent.py 第一行的 `import langchain.agents` 就会带入它，故必须放在这里。）
- 若以后要下载新模型：注释掉 HF_HUB_OFFLINE 一行，即自动走 HF_ENDPOINT 国内镜像。
"""

import os

os.environ.setdefault("HF_ENDPOINT", "https://hf-mirror.com")
os.environ.setdefault("HF_HUB_OFFLINE", "1")
# 关闭 Chroma 匿名遥测：离线环境上报必然失败，只会刷无害报错日志
os.environ.setdefault("ANONYMIZED_TELEMETRY", "False")
