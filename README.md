
# 📦 项目名称：电子元器件 RAG 智能问答系统

本项目构建了一个端到端的文档级 **RAG（Retrieval-Augmented Generation）问答系统**，用于结构化和非结构化电子元器件资料（Excel/PPT/Word）智能问答，支持本地索引、语义检索、DeepSeek 生成式回答。

---

## 🚀 功能概述

| 模块 | 描述 |
|------|------|
| 📄 `excel_parser.py` | 将 Excel 文件结构化为 JSON，支持合并单元格展开与按字符数分块 |
| 📊 `pptx_parser.py` | 解析 PPT 文件，每页视为一个 sheet，提取文本和表格数据 |
| 📚 `word_parser.py` | 解析 Word 文件段落与表格内容，统一转换为 JSON 格式 |
| 🤖 `excel_llm_main.py` | 使用 Gemini API 对结构化 JSON 内容进行规范化与错误修正 |
| 🧱 `save_to_mysql.py` | 将 LLM 处理后的 JSON 写入 MySQL 数据库（含 doc_id 分配） |
| 🔍 `save_to_es.py` | 从 MySQL 批量导入至 Elasticsearch，支持语义检索与上下文拼接 |
| 🧠 `rag_with_deepseek.py` | 基于 DeepSeek-R1 模型执行文档级 RAG 问答 |
| 🧪 `es_main.py` | 用于调试索引创建、清空、测试检索、打印结果 |
| ⚙️ `gemini_support.py` | 调试 Gemini API 可用模型列表 |

---

## 🧩 系统流程图

```text
        ┌─────────────┐
        │ Excel/PPT/Word 原始文件 │
        └─────┬───────┘
              ▼
    ┌────────────────────┐
    │ 解析为结构化 JSON   │ ← excel_parser / pptx_parser / word_parser
    └────────┬───────────┘
             ▼
     ┌─────────────────────┐
     │ Gemini LLM 校对美化 │ ← excel_llm_main.py
     └────────┬────────────┘
              ▼
     ┌──────────────────────┐
     │ MySQL 写入 + 去重管理 │ ← save_to_mysql.py
     └────────┬─────────────┘
              ▼
     ┌────────────────────────────┐
     │ ES 建索引+多字段语义检索    │ ← save_to_es.py
     └────────┬──────────────────┘
              ▼
        ┌─────────────────────┐
        │ DeepSeek-R1 回答生成 │ ← rag_with_deepseek.py
        └─────────────────────┘
```

---

## 📁 数据结构说明

- 所有文档将被转换为以下 JSON 格式：
```json
{
  "doc_type": "excel | ppt | word",
  "file_name": "xxx.xlsx",
  "tables": [
    {
      "sheet": "sheet名称或标题",
      "data": "原始CSV文本",
      "rows": [ { "列名": "值", ... } ],
      "text": "（PPT或Word才有）"
    }
  ]
}
```

---

## ⚙️ 环境依赖

- Python 3.9+
- `elasticsearch`
- `mysql-connector-python`
- `openpyxl`
- `python-docx`
- `pandas`
- `google-generativeai`（Gemini）
- `volcenginesdkarkruntime`（DeepSeek）

---

## 🔑 环境变量（.env）

```
GEMINI_API_KEY=你的Gemini密钥
VOLCENGINE_API_KEY=你的火山引擎密钥
VOLCENGINE_ENDPOINT_ID=DeepSeek R1 推理接入点ID
```

---

## 🛠️ 快速开始

```bash
# 1. Excel/PPT/Word 转换为 JSON
python excel_parser.py

# 2. Gemini 模型修正
python excel_llm_main.py

# 3. 存入 MySQL
python save_to_mysql.py

# 4. 写入 Elasticsearch
python es_main.py

# 5. 执行 RAG 查询
python rag_with_deepseek.py
```

---

## 📌 备注

- MySQL 和 Elasticsearch 地址默认配置为局域网 IP，如需更改请修改 `save_to_es.py` 和 `save_to_mysql.py` 中连接参数。
- 所有 JSON 输出默认存储在 `llm_output_test/` 或 `output_test/` 目录。
