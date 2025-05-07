from volcenginesdkarkruntime import Ark
import os
from save_to_es import Elastic

# 配置火山引擎 API
VOLCENGINE_API_KEY = os.getenv("VOLCENGINE_API_KEY")  # 从环境变量读取火山引擎 API 密钥
VOLCENGINE_ENDPOINT_ID = os.getenv("VOLCENGINE_ENDPOINT_ID")  # 从环境变量读取推理接入点 ID

# 初始化 Ark 客户端
client = Ark(api_key=VOLCENGINE_API_KEY)

def generate_with_deepseek(query, context):
    """
    使用 Ark 客户端调用 DeepSeek R1 模型生成结果，适配 RAG 任务
    """
    if not VOLCENGINE_API_KEY:
        raise ValueError("未设置 VOLCENGINE_API_KEY 环境变量")
    if not VOLCENGINE_ENDPOINT_ID:
        raise ValueError("未设置 VOLCENGINE_ENDPOINT_ID 环境变量")


    # System 提示：明确任务和要求
    system_prompt = (
        "你是问答任务助手，擅长从给定的上下文中提取信息并回答问题。你的回答必须以上下文内容为主要依据，优先关注与问题最相关的内容，直接引用相关信息。如果需要简单推理，推理必须基于上下文内容，禁止引入上下文之外的常识或外部知识。如果上下文不足以回答问题，直接回答“未找到相关答案”。\n\n"

    )

    # User 提示：整合上下文和问题
    user_prompt = (
        f"以下是上下文：\n"
        f"{context}\n\n"
        f"问题：\n"
        f"{query}\n\n"
        "请根据上下文回答上述问题，优先引用上下文中的具体内容，并做出详细解答："
    )

    try:
        response = client.chat.completions.create(
            model=VOLCENGINE_ENDPOINT_ID,  # 使用推理接入点 ID
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            stream=False
        )

        return response.choices[0].message.content
    except Exception as e:
        print(f"DeepSeek API 调用失败: {e}")
        return "生成失败"

def rag_pipeline(query, index_name="e_rag"):
    """
    RAG 完整流程：从 ES 查询到生成结果，并返回检索到的文档名称
    返回：
    - generated_text: LLM 生成的回答
    - doc_sources: 检索到的文档来源列表 [(file_name, sheet_name), ...]
    """
    # 初始化 ES 客户端
    es = Elastic()

    # 步骤 1：搜索并构建 context，同时获取文档来源
    context, doc_sources = es.search_and_build_context(index_name, query)
    if context.startswith("未找到"):
        return context, []

    # 步骤 2：调用 DeepSeek R1 生成
    generated_text = generate_with_deepseek(query, context)
    return generated_text, doc_sources

# 示例使用
if __name__ == "__main__":
    query = "抹布烘干的事项有哪些"
    generated_text, doc_sources = rag_pipeline(query, index_name="e_rag")
    
    # 准备输出内容
    output_lines = []
    
    # 添加检索到的文档名称
    output_lines.append("检索到的文档名称：")
    if doc_sources:
        for file_name, sheet_name in doc_sources:
            line = f"- {file_name}_{sheet_name}"
            output_lines.append(line)
    else:
        output_lines.append("未检索到相关文档。")
    
    # 添加生成结果
    output_lines.append("\n生成结果：")
    output_lines.append(generated_text)
    
    # 打印到控制台
    for line in output_lines:
        print(line)
    
    # 保存到文件
    output_file = "output.txt"
    try:
        with open(output_file, "w", encoding="utf-8") as f:
            f.write("\n".join(output_lines))
        print(f"\n结果已保存到 {output_file}")
    except Exception as e:
        print(f"保存文件失败: {e}")