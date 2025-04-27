import google.generativeai as genai
import os

# 配置 API 密钥（确保已从 Google AI Studio 获取）
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

# 列出所有可用模型
models = genai.list_models()
for model in models:
    print(model.name)