from markitdown import MarkItDown
from openai import OpenAI
from dotenv import load_dotenv
import os

# 加载 .env 文件
load_dotenv()

# 确认 API 密钥已加载
api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    raise ValueError("OPENAI_API_KEY not found in .env file")

# 初始化 OpenAI 客户端
client = OpenAI(api_key=api_key)  # 显式传递 API 密钥

# 初始化 MarkItDown，启用 LLM
file_path = r"C:\Users\dreame\Desktop\电子元件RAG\数据表格纯文字+复杂图片\IMU选型参数对比表.xlsx"
md = MarkItDown(llm_client=client, llm_model="gpt-4o", enable_plugins=False)

# 转换 Excel 文件
try:
    result = md.convert(file_path)
    # 保存结果到文件
    output_file = r"C:\Users\dreame\Desktop\电子元件RAG\output_with_images.txt"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(result.text_content)
    print(f"Content has been saved to {output_file}")
except Exception as e:
    print(f"Error occurred: {e}")