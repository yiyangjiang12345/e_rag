from save_to_es import Elastic

def main():
    
    es = Elastic()
    #print(es.create_label_index("e_rag"))

    # 清空现有文档（可选）
    #print(es.clear_documents("e_rag"))

    #print(es.bulk_index_data("e_rag", database="e_rag"))
    # 搜索关键词
    file_names, sheet_names, json_contents = es.search_by_text("e_rag", "自研风机")
    # 组合 file_names 和 sheet_names 为 file_name_sheet_name 格式
    combined_names = [f"{file_name}_{sheet_name}" for file_name, sheet_name in zip(file_names, sheet_names)]
    
    print("搜索结果：")
    print("文件名:", combined_names)
    print("JSON内容:", json_contents)

if __name__ == "__main__":
    main()
