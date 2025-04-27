# -*- coding: utf-8 -*-
import os
import mysql.connector
from mysql.connector import Error

def connect_to_mysql():
    """
    连接到 MySQL 数据库。
    :return: 数据库连接对象
    """
    try:
        connection = mysql.connector.connect(
            host="10.10.37.77",
            user="root",
            password="TF123456",
            database="excel_data",
            charset="utf8mb4"
        )
        if connection.is_connected():
            print("Successfully connected to MySQL database")
        return connection
    except Error as e:
        print(f"Error connecting to MySQL: {str(e)}")
        raise

def save_to_mysql(connection, doc_id: str, file_name: str, sheet_name: str, json_content: str):
    """
    将 JSON 数据保存到 MySQL 数据库，先检查是否存在重复记录，避免自增 id 跳跃。
    :param connection: 数据库连接对象
    :param doc_id: 文档 ID（根据 file_name 生成）
    :param file_name: Excel 文件名
    :param sheet_name: Sheet 名称
    :param json_content: JSON 字符串
    """
    cursor = None
    try:
        cursor = connection.cursor()
        
        # 检查是否已存在相同 file_name 和 sheet_name 的记录
        check_query = """
        SELECT COUNT(*) FROM llm_outputs WHERE file_name = %s AND sheet_name = %s
        """
        cursor.execute(check_query, (file_name, sheet_name))
        count = cursor.fetchone()[0]
        
        if count > 0:
            print(f"Duplicate record found for {file_name}_{sheet_name}, skipping insertion...")
            return  # 直接返回，不执行插入操作，避免自增 id 分配
        
        # 如果记录不存在，执行插入
        insert_query = """
        INSERT INTO llm_outputs (doc_id, file_name, sheet_name, json_content)
        VALUES (%s, %s, %s, %s)
        """
        data = (doc_id, file_name, sheet_name, json_content)
        cursor.execute(insert_query, data)
        connection.commit()
        print(f"Successfully saved data to MySQL: {file_name}_{sheet_name}")
        
    except Error as e:
        connection.rollback()
        print(f"Error saving to MySQL: {str(e)}")
        raise
    finally:
        if cursor:
            cursor.close()

def extract_info_from_filename(filename: str) -> tuple[str, str]:
    """
    从文件名中提取 file_name 和 sheet_name。
    文件名格式：
    - Excel 文件：excel文件名_sheet名称_llm_output_0.json
    - Docx 文件：word名称_1.json（无 sheet_name，使用 file_name 作为 sheet_name）
    - Pptx 文件：ppt名称_sheet名称_2.json
    :param filename: 文件名（不含路径）
    :return: (file_name, sheet_name)
    """
    # 移除后缀 .json
    base_name = filename.replace(".json", "")
    
    # 检查文件名是否以 _llm_output_0, _1, 或 _2 结尾
    if base_name.endswith("_llm_output_0"):
        # Excel 文件：excel文件名_sheet名称_llm_output_0
        type_suffix = "llm_output_0"
        file_part = base_name[:-len("_llm_output_0")]  # 移除 _llm_output_0
        # 按最后一个 _ 分割，提取 excel文件名 和 sheet名称
        sub_parts = file_part.rsplit("_", 1)
        if len(sub_parts) != 2:
            raise ValueError(f"Excel filename format invalid: {filename}. Expected format: excel文件名_sheet名称_llm_output_0.json")
        file_name_base = sub_parts[0]
        sheet_name = sub_parts[1]
        file_name = f"{file_name_base}.xlsx"
    elif base_name.endswith("_1"):
        # Docx 文件：word名称_1
        type_suffix = "1"
        file_part = base_name[:-len("_1")]  # 移除 _1
        file_name_base = file_part
        file_name = f"{file_name_base}.docx"
        sheet_name = file_name  # 使用 file_name 作为 sheet_name
    elif base_name.endswith("_2"):
        # Pptx 文件：ppt名称_sheet名称_2
        type_suffix = "2"
        file_part = base_name[:-len("_2")]  # 移除 _2
        # 按最后一个 _ 分割，提取 ppt名称 和 sheet名称
        sub_parts = file_part.rsplit("_", 1)
        if len(sub_parts) != 2:
            raise ValueError(f"Pptx filename format invalid: {filename}. Expected format: ppt名称_sheet名称_2.json")
        file_name_base = sub_parts[0]
        sheet_name = sub_parts[1]
        file_name = f"{file_name_base}.pptx"
    else:
        raise ValueError(f"Invalid file type suffix in filename {filename}. Expected suffix: _llm_output_0, _1, or _2")

    return file_name, sheet_name

def main():
    # 定义输入目录（大模型处理后的 JSON 文件）
    input_dir = "llm_output_test"  # 存储大模型校对后文件的目录

    # 确保输入目录存在
    if not os.path.exists(input_dir):
        raise FileNotFoundError(f"Input directory not found: {input_dir}")

    # 连接到 MySQL 数据库
    connection = connect_to_mysql()

    # 用于跟踪 file_name 到 doc_id 的映射
    file_name_to_doc_id = {}
    next_doc_id = 0  # 从 0 开始自增

    try:
        # 遍历 llm_output 目录中的所有 JSON 文件
        for json_file in os.listdir(input_dir):
            # 适配不同文件类型：_llm_output_0.json（Excel），_1.json（Docx），_2.json（Pptx）
            if not (json_file.endswith("_llm_output_0.json") or json_file.endswith("_1.json") or json_file.endswith("_2.json")):
                continue

            input_json_path = os.path.join(input_dir, json_file)
            print(f"\n=== Processing JSON file: {input_json_path} ===")

            # 从文件名中提取 file_name 和 sheet_name
            try:
                file_name, sheet_name = extract_info_from_filename(json_file)
            except ValueError as e:
                print(f"Error processing filename {json_file}: {str(e)}")
                continue

            # 去掉 file_name 的后缀，用于映射 doc_id
            file_name_base = os.path.splitext(file_name)[0]

            # 为 file_name 分配 doc_id
            if file_name_base not in file_name_to_doc_id:
                file_name_to_doc_id[file_name_base] = str(next_doc_id)
                next_doc_id += 1
            doc_id = file_name_to_doc_id[file_name_base]
            print(f"Assigned doc_id: {doc_id} for file_name: {file_name}")

            # 读取 JSON 文件内容（作为字符串）
            with open(input_json_path, "r", encoding="utf-8") as f:
                json_str = f.read()

            # 调试：检查 JSON 字符串的完整性
            print(f"JSON string length: {len(json_str)} characters")
            print(f"First 500 characters:\n{json_str[:500]}")
            print(f"Last 500 characters:\n{json_str[-500:]}")

            # 保存到 MySQL 数据库
            save_to_mysql(connection, doc_id, file_name, sheet_name, json_str)

    finally:
        # 关闭数据库连接
        if connection.is_connected():
            connection.close()
            print("MySQL connection closed")

if __name__ == "__main__":
    main()