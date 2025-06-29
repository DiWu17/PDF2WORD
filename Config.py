# -- Config.py --
# 全局配置文件

import os

# ==============================================================================
# 大模型 API 配置 (VLM for Table Recognition)
# ==============================================================================
# 本项目使用阿里云DashScope（通义千问）的VLM（视觉语言模型）来识别和转换图片中的表格。
# 您需要在这里配置您的API密钥和基础URL。
#
# 获取API Key: https://help.aliyun.com/zh/dashscope/developer-reference/activate-dashscope-and-create-an-api-key
# 
# 请将 "your_real_api_key" 替换为您的真实API密钥。
# 推荐使用环境变量来管理密钥，以增强安全性。
# 例如: os.getenv("DASHSCOPE_API_KEY", "your_default_key_if_not_set")

DASHSCOPE_API_KEY = ""
# DashScope API的基础URL，通常保持默认即可。
DASHSCOPE_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"

# 是否使用流式请求，默认为True，以优化性能和用户体验。
# 如果设置为False，则会等待模型完全生成所有内容后一次性返回，可能会增加等待时间。
DASHSCOPE_STREAM_REQUEST = False
