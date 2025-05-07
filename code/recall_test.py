import pandas as pd
import logging
from save_to_es import Elastic  # 假设 Elastic 类在 Elastic.py 文件中

# 配置日志
logging.basicConfig(filename='recall_test.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

def load_excel_data(excel_file):
    """
    从 Excel 文件加载测试数据集，提取问题和预期文档。
    参数：
        excel_file: Excel 文件路径
    返回：
        数据列表，格式为 [{'问题': str, '文档名_表名': str}, ...]
    """
    # 读取 Excel 文件
    df = pd.read_excel(excel_file)
    
    # 确保列名正确
    expected_columns = ['问题', '文档名_表名']
    if not all(col in df.columns for col in expected_columns):
        raise ValueError(f"Excel 文件缺少必要列：{expected_columns}")

    # 提取问题和文档名_表名
    test_data = df[['问题', '文档名_表名']].to_dict('records')
    
    logging.info(f"Loaded {len(test_data)} queries from {excel_file}")
    return test_data

def calculate_recall(es, index_name, test_data):
    """
    计算 Recall@1, Recall@3, Recall@5。
    参数：
        es: Elastic 实例
        index_name: 索引名称
        test_data: 测试数据集（从 Excel 加载）
    返回：
        recall_at_n: 包含 Recall@1, Recall@3, Recall@5 的字典
    """
    total_queries = len(test_data)
    hits_at_1 = 0
    hits_at_3 = 0
    hits_at_5 = 0

    for row in test_data:
        query = row['问题']
        expected_file_name, expected_sheet_name = row['文档名_表名'].split('_')
        
        # 执行搜索
        try:
            file_names, sheet_names, json_contents, scores = es.search_by_text(index_name, query)
        except Exception as e:
            logging.error(f"Query failed: {query}, Error: {str(e)}")
            continue
        
        # 检查前 n 条结果是否包含正确文档
        for i, (fn, sn) in enumerate(zip(file_names[:5], sheet_names[:5])):
            if fn == expected_file_name and sn == expected_sheet_name:
                if i < 1:
                    hits_at_1 += 1
                    hits_at_3 += 1
                    hits_at_5 += 1
                elif i < 3:
                    hits_at_3 += 1
                    hits_at_5 += 1
                elif i < 5:
                    hits_at_5 += 1
                logging.info(f"Query: {query}, Hit at rank {i+1}, Expected: {expected_file_name}_{expected_sheet_name}")
                break
        else:
            logging.warning(f"Query: {query}, No hit in top 5, Expected: {expected_file_name}_{expected_sheet_name}")

    recall_at_1 = hits_at_1 / total_queries if total_queries > 0 else 0
    recall_at_3 = hits_at_3 / total_queries if total_queries > 0 else 0
    recall_at_5 = hits_at_5 / total_queries if total_queries > 0 else 0

    return {
        'Recall@1': recall_at_1,
        'Recall@3': recall_at_3,
        'Recall@5': recall_at_5
    }

def main():
    """
    主函数：加载 Excel 数据、执行召回检索、计算召回率。
    """
    # 初始化 Elastic 实例
    es = Elastic()
    index_name = "e_rag"  # 请确认实际索引名称
    excel_file = "合并问答测试数据集_完整版80条.xlsx"

    # 加载 Excel 数据
    try:
        test_data = load_excel_data(excel_file)
    except FileNotFoundError:
        print(f"错误：未找到 {excel_file}")
        return
    except ValueError as e:
        print(f"错误：{str(e)}")
        return

    # 计算召回率
    recall_results = calculate_recall(es, index_name, test_data)
    
    # 输出结果
    print("\n召回率测试结果：")
    for n, recall in recall_results.items():
        print(f"{n}: {recall:.4f}")
    
    logging.info(f"Recall results: {recall_results}")

if __name__ == "__main__":
    main()
