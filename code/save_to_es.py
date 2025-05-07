from elasticsearch import Elasticsearch
from elasticsearch.helpers import bulk
import mysql.connector
import json

class Elastic(object):
    def __init__(self, hosts="http://10.10.37.75:9200"):
        self.client = Elasticsearch(hosts=hosts)

    def get(self, name, id):
        source = ["json_content", "sheet_name", "file_name"]
        return self.client.get(index=name, id=id, _source=source)

    def create_label_index(self, name, number_of_replicas=0, number_of_shards=1):
        """
        调整映射以适配你的数据结构：
        - file_name: 文件名
        - sheet_name: 工作表名
        - json_content: JSON内容，改为text类型，支持检索
        """
        # 检查索引是否已经存在
        if self.client.indices.exists(index=name):
            print(f"Index '{name}' already exists, skipping creation.")
            return "索引已存在，跳过创建"

        mappings = {
            "properties": {
                "file_name": {
                    "type": "text",
                    "analyzer": "ik_max_word",  
                    "search_analyzer": "ik_smart",
                },
                "sheet_name": {
                    "type": "text",
                    "analyzer": "ik_max_word",
                    "search_analyzer": "ik_smart",
                },
                "json_content": {
                    "type": "text",  
                    "analyzer": "ik_max_word",
                    "search_analyzer": "ik_smart",
                },
            }
        }
        setting = {
            "settings": {
                "number_of_replicas": number_of_replicas,
                "number_of_shards": number_of_shards,
            },
            "mappings": mappings,
        }
        self.client.indices.create(index=name, body=setting)
        return "创建索引成功"

    def clear_documents(self, name):
        """
        清空索引中的所有文档，但保留索引结构。
        """
        if self.client.indices.exists(index=name):
            self.client.delete_by_query(
                index=name,
                body={
                    "query": {
                        "match_all": {}
                    }
                }
            )
            print(f"All documents in index '{name}' deleted.")
        else:
            print(f"Index '{name}' does not exist.")
        return "清空文档完成"

    def bulk_index_data(
        self,
        name,
        database="e_rag",  
        batch_size=64,
    ):
        """
        从MySQL中读取数据并批量插入到ES
        MySQL连接信息：
        - 地址：10.10.37.77
        - 账号：root
        - 密码：TF123456
        """
        # MySQL连接配置
        mysql_config = {
            "host": "10.10.37.77",
            "user": "root",
            "password": "TF123456",
            "database": database
        }

        # 连接MySQL
        conn = mysql.connector.connect(**mysql_config)
        cursor = conn.cursor(dictionary=True)  # 返回字典格式的结果

        # 查询数据
        query = "SELECT id, file_name, sheet_name, json_content FROM llm_outputs;"
        cursor.execute(query)
        
        # 分批处理
        while True:
            rows = cursor.fetchmany(batch_size)  # 每次获取batch_size条数据
            if not rows:  # 如果没有更多数据，退出循环
                break

            requests = []
            for row in rows:
                # json_content可能存储为字符串，需要解析
                json_content_str = row["json_content"]
                try:
                    # 如果json_content是字符串，尝试解析为JSON并转为字符串形式用于检索
                    json_content = json.dumps(json.loads(json_content_str), ensure_ascii=False)
                except (json.JSONDecodeError, TypeError):
                    json_content = json_content_str  # 如果解析失败，直接使用原始字符串

                request = {
                    "_op_type": "index",
                    "_index": name,
                    "_id": row["id"],
                    "_source": {
                        "file_name": row["file_name"],
                        "sheet_name": row["sheet_name"],
                        "json_content": json_content,
                    },
                }
                requests.append(request)

            # 批量插入到ES
            bulk(self.client, requests)

        # 关闭MySQL连接
        cursor.close()
        conn.close()
        return "插入数据成功"

    def search_by_text(self, name, text):
        dsl_text = {
            "_source": ["file_name", "sheet_name", "json_content"],
            "size": 10,
            "query": {
                "match": {
                    "json_content": text  # 只在json_content中检索
                }
            },
        }
        result = self.client.search(index=name, body=dsl_text)
        hits = result["hits"]["hits"]
        file_names = [x["_source"]["file_name"] for x in hits]
        sheet_names = [x["_source"]["sheet_name"] for x in hits]
        json_contents = [x["_source"]["json_content"] for x in hits]
        scores = [x["_score"] for x in hits]
        return file_names, sheet_names, json_contents, scores

    def search_and_build_context(self, name, text):
        """
        搜索并构建上下文，用于 RAG 输入
        参数：
        - name: ES 索引名称
        - text: 查询关键词
        返回：
        - context: 构建好的上下文字符串
        - doc_sources: 检索到的文档来源列表 [(file_name, sheet_name), ...]
        """
        # 步骤 1：从 ES 查询
        file_names, sheet_names, json_contents, scores = self.search_by_text(name, text)
        if not file_names:
            return "未找到相关内容", []

        # 步骤 2：按得分排序
        results = [
            {"file_name": fn, "sheet_name": sn, "content": jc, "score": s}
            for fn, sn, jc, s in zip(file_names, sheet_names, json_contents, scores)
        ]
        results.sort(key=lambda x: x["score"], reverse=True)

        # 步骤 3：动态召回
        selected_results = []
        doc_sources = []  # 存储文档来源
        total_length = 0
        min_length = 50000  # 最小长度
        max_length = 60000  # 最大长度

        # 默认取前 5 个结果
        for i in range(min(5, len(results))):
            result = results[i]
            part = f"[来源: {result['file_name']}_{result['sheet_name']}]\n{result['content']}"
            selected_results.append(part)
            doc_sources.append((result['file_name'], result['sheet_name']))
            total_length += len(part)

        # 如果总长度 < 50000，继续取下一个，直到接近 60000
        if total_length < min_length:
            for i in range(5, len(results)):
                result = results[i]
                part = f"[来源: {result['file_name']}_{result['sheet_name']}]\n{result['content']}"
                part_length = len(part)

                if total_length + part_length > max_length:
                    remaining_length = max_length - total_length
                    if remaining_length > 0:
                        part = part[:remaining_length] + "..."
                        selected_results.append(part)
                        doc_sources.append((result['file_name'], result['sheet_name']))
                    break

                selected_results.append(part)
                doc_sources.append((result['file_name'], result['sheet_name']))
                total_length += part_length

        # 步骤 4：用分隔符拼接
        context = "   ---   ".join(selected_results)
        return context, doc_sources