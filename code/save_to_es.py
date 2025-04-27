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
        database="excel_data",  
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
        file_names = [x["_source"]["file_name"] for x in result["hits"]["hits"]]
        sheet_names = [x["_source"]["sheet_name"] for x in result["hits"]["hits"]]
        json_contents = [x["_source"]["json_content"] for x in result["hits"]["hits"]]
        return file_names, sheet_names, json_contents