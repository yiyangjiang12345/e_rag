# -*- coding: utf-8 -*-
import json
from datetime import datetime
from typing import Dict, Any
import google.generativeai as genai
from dotenv import load_dotenv
import os
import time
import re  # 用于解析文件名中的 chunk 编号

# 加载 .env 文件
load_dotenv()

def correct_json_with_gemini(input_json: Dict[str, Any], retries=3) -> str:
    """
    调用 Gemini API，使用输入的 JSON 数据和提示词，生成校对后的 JSON 字符串。
    （不再强制用 json.loads 验证）
    :param input_json: 读取的 JSON 数据（单个 sheet）
    :param retries: 重试次数
    :return: 校对后的“JSON字符串”文本（不保证一定是有效JSON）
    """
    # 从环境变量中获取 API 密钥
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY not found in .env file")

    # 配置 Gemini API
    genai.configure(api_key=api_key)

    # 将输入的 JSON 数据转换为字符串（紧凑格式以减少 token 数量）
    json_str = json.dumps(input_json, ensure_ascii=False, indent=None, separators=(",", ":"))

    # 调试：检查 json_str 的完整性
    print("\n=== JSON String Before Prompt ===")
    print(f"JSON string length: {len(json_str)} characters")
    print(f"First 500 characters:\n{json_str[:500]}")
    print(f"Last 500 characters:\n{json_str[-500:]}")

    # 创建提示词，嵌入 JSON 数据
    prompt = (
        "你是一个JSON数据校正和标准化专家。我将提供一个从Excel文件解析得到的JSON对象，其中包含表格数据（tables），但可能混杂自然段内容。你的任务是校对和修正这个JSON数据，确保其符合正确的JSON表格形式，同时最大程度保留自然段的格式。JSON数据可能较大，请确保完整处理所有内容。具体要求如下："
        "1.确保所有字段名和值都保持中文，不要将中文转换为英文。"
        "2.识别表格数据和自然段内容，并保持格式正确"
        "3.移除冗余字段（例如key和value值都为默认或者空的字段，如“Column_1: ''”），但保留有意义的空值（例如表格中的空单元格，如“ST: ''”）。"
        "4.如果JSON数据较大，请分块处理，确保不遗漏任何内容。"
        "5.确保输出的JSON结构清晰、格式正确，所有内容均为中文。"
        "6.只返回校对后的有效JSON字符串，不要包含其他任何文字或注释，不要包含任何前后缀（如```json或'''）。"
        "以下是输入的JSON数据："
        f"{json_str}"
        "请返回校对后的JSON字符串。"
    )

    # 调试：检查 prompt 的完整性
    print("\n=== Prompt Before API Call ===")
    print(f"Prompt length: {len(prompt)} characters")
    print(f"First 500 characters of prompt:\n{prompt[:500]}")
    print(f"Last 500 characters of prompt:\n{prompt[-500:]}")

    for attempt in range(retries):
        try:
            # 调用 Gemini API
            model = genai.GenerativeModel('models/gemini-2.5-pro-preview-03-25')  # 更新为可用模型
            response = model.generate_content(prompt)

            # 获取校对后的 JSON 字符串（此处只当作文本，不再做 json.loads 校验）
            corrected_json_str = response.text.strip()

            # 移除可能的 ```json 前缀和 ``` 后缀，或者 ''' 前后缀
            corrected_json_str = corrected_json_str.strip()
            if corrected_json_str.startswith("```json"):
                corrected_json_str = corrected_json_str[len("```json"):].strip()
            if corrected_json_str.endswith("```"):
                corrected_json_str = corrected_json_str[:-len("```")].strip()
            if corrected_json_str.startswith("'''"):
                corrected_json_str = corrected_json_str[len("'''"):].strip()
            if corrected_json_str.endswith("'''"):
                corrected_json_str = corrected_json_str[:-len("'''")].strip()

            print("\n原始 API 返回的内容（清理前后缀后）：")
            print(corrected_json_str)

            return corrected_json_str

        except Exception as e:
            if attempt == retries - 1:
                raise Exception(f"Error calling Gemini API after {retries} attempts: {str(e)}")
            print(f"Attempt {attempt + 1} failed, retrying... Error: {str(e)}")
            time.sleep(2)

def main():
    # 定义输入和输出目录
    input_dir = "output_test"  # 存储原始 JSON 文件的目录
    output_dir = "llm_output_test"  # 存储大模型校对后文件的目录

    # 字符数限制
    CHARACTER_LIMIT = 70000  # 字符数超过 70000 时跳过大模型处理

    # 确保输入目录存在
    if not os.path.exists(input_dir):
        raise FileNotFoundError(f"Input directory not found: {input_dir}")

    # 确保输出目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 遍历 input_dir 中的所有 JSON 文件
    for json_file in os.listdir(input_dir):
        if not json_file.endswith(".json"):
            continue

        input_json_path = os.path.join(input_dir, json_file)
        print(f"\n=== Processing JSON file: {input_json_path} ===")

        # 读取 JSON 文件
        with open(input_json_path, "r", encoding="utf-8") as f:
            parsed_json = json.load(f)

        # 调试：检查 parsed_json 的完整性
        print(f"Total tables: {len(parsed_json['tables'])}")
        for i, table in enumerate(parsed_json['tables']):
            print(f"Table {i + 1} - Sheet: {table['sheet']}")
            print(f"Number of rows: {len(table['rows'])}")
            print(f"Data length: {len(table['data'])} characters")

        # 将 JSON 数据转换为字符串（紧凑格式）
        json_str = json.dumps(parsed_json, ensure_ascii=False, indent=None, separators=(",", ":"))

        # 检查字符数
        json_str_length = len(json_str)
        print(f"JSON string length: {json_str_length} characters")

        if json_str_length > CHARACTER_LIMIT:
            print(f"JSON string exceeds {CHARACTER_LIMIT} characters, skipping Gemini API processing...")
            corrected_json_str = json_str  # 直接使用原始 JSON 字符串
        else:
            # 调用大模型校对
            print("Processing with Gemini API...")
            corrected_json_str = correct_json_with_gemini(parsed_json)

        # 从输入文件名中提取 excel 文件名、sheet 名称和 chunk 编号
        # 去掉扩展名，例如 "R2350电机保护逻辑_Sheet11_0.json" -> "R2350电机保护逻辑_Sheet11_0"
        base_file_name = os.path.splitext(json_file)[0]
        # 以 "_" 分割，例如 ["R2350电机保护逻辑", "Sheet11", "0"]
        parts = base_file_name.split("_")
        if len(parts) < 3:
            raise ValueError(f"Invalid input file name format: {json_file}")

        # 提取 excel 文件名（可能是多个部分，例如 "R2350电机保护逻辑"）
        excel_name = "_".join(parts[:-2])  # 取除最后两个部分之前的所有部分
        # 提取 sheet 名称和可能的 chunk 编号
        sheet_part = parts[-2]  # 倒数第二部分，例如 "Sheet11" 或 "Sheet2"

        # 使用正则表达式分离 sheet 名称和 chunk 编号
        match = re.match(r"^(.*?)([0-9]+)$", sheet_part)
        if match:
            # 分块文件，例如 "Sheet11" -> sheet_name = "Sheet1", chunk_number = "1"
            sheet_name = match.group(1)  # "Sheet1"
            chunk_number = match.group(2)  # "1"
            output_file_name = f"{excel_name}_{sheet_name}{chunk_number}_llm_output_0.json"
        else:
            # 非分块文件，例如 "Sheet2" -> sheet_name = "Sheet2"
            sheet_name = sheet_part
            output_file_name = f"{excel_name}_{sheet_name}_llm_output_0.json"

        # 构造输出路径
        output_path = os.path.join(output_dir, output_file_name)

        # 保存校对后的 JSON 字符串到文件
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(corrected_json_str)
        print(f"\nSaved JSON to: {output_path}")

if __name__ == "__main__":
    main()