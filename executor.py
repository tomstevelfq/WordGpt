from word_api import *
from llms import llm
from langchain.prompts import PromptTemplate
from langchain.chains.llm import LLMChain
import json

# 启动word
start_word()

excutor_prompt="""
您将收到一个任务指令，任务指令描述了您需要执行的任务。如果任务指令是结束，你就返回结束就行，任务可以通过以下 API 进行解决：

1. 启动 Word 程序：`start_word()`
2. 打开现有文档：`open_document(file_path)`
3. 创建新文档：`create_document()`
4. 保存文档：`save_document(file_path)`
5. 关闭文档：`close_document()`
6. 退出 Word 程序：`quit_word()`
7. 添加段落并设置格式：`add_paragraph(text, style="正文", font_name="宋体", font_size=12, bold=False, italic=False, alignment=1)`
8. 设置文档标题：`set_title(title_text, style="标题 1", font_name="宋体", font_size=18)`
9. 设置段落对齐方式：`set_paragraph_alignment(paragraph_index, alignment)`
10. 设置段落字体和字号：`set_paragraph_font(paragraph_index, font_name, font_size)`
11. 设置段落加粗：`set_paragraph_bold(paragraph_index, bold)`
12. 设置段落斜体：`set_paragraph_italic(paragraph_index_list, italic)`
13. 修改段落样式：`modify_paragraph_style(paragraph_index, style="正文", font_name="宋体", font_size=12)`
14. 修改段落文本：`modify_paragraph(paragraph_index, new_text)`
15. 获取 Word 文档内容：`get_word_content()`

任务指令：{task_instruction}

根据任务指令，您需要给出相应的 API 调用。API 调用应包括以下内容：
1. API 名称
2. API 调用的参数信息
3. API 调用的格式，采用 JSON 格式给出

示例：
任务指令：为文档添加标题并设置字体为宋体，字体大小为18。
API 调用：
{
  "api_name": "set_title",
  "params": {
    "title_text": "文档标题",
    "style": "标题 1",
    "font_name": "宋体",
    "font_size": 18
  }
}

任务指令：修改第一段文本内容为“这是新的段落”，并设置为加粗。
API 调用：
{
  "api_name": "modify_paragraph",
  "params": {
    "paragraph_index": 1,
    "new_text": "这是新的段落"
  }
}

任务指令：获取当前文档内容并转换为 JSON 格式。
API 调用：
{
  "api_name": "get_word_content",
  "params": {}
}
请根据您收到的任务指令，按照上述格式给出对应的 API 调用。
"""

# 执行任务的函数
def execute_task(task_instruction: str) -> dict:
    """
    根据任务指令生成相应的 API 调用结构体
    :param task_instruction: 任务指令
    :param llm: 使用的 LLM 模型
    :return: API 调用结构体
    """
    
    # 使用 LLMChain 调用大模型进行推理
    prompt = PromptTemplate(
        template=excutor_prompt,
        input_variables=["task_instruction"]
    )
    chain = LLMChain(llm=llm, prompt=prompt)
    
    # 执行任务指令并获取模型输出
    output = chain.run(task_instruction=task_instruction)
    
    # 从模型输出中获取 API 调用结构体
    # 输出应该是一个 JSON 格式的字符串
    try:
        api_call = json.loads(output)
    except json.JSONDecodeError:
        raise ValueError("无法解析 API 调用格式，返回内容不符合预期")

    return api_call

# 将所有API存储在字典中
api_functions = {
    "start_word": start_word,
    "open_document": open_document,
    "create_document": create_document,
    "save_document": save_document,
    "close_document": close_document,
    "quit_word": quit_word,
    "add_paragraph": add_paragraph,
    "set_title": set_title,
    "set_paragraph_alignment": set_paragraph_alignment,
    "set_paragraph_font": set_paragraph_font,
    "set_paragraph_bold": set_paragraph_bold,
    "set_paragraph_italic": set_paragraph_italic,
    "modify_paragraph_style": modify_paragraph_style,
    "modify_paragraph": modify_paragraph,
    "get_word_content": get_word_content
}

# 执行API调用的函数
def execute_api_call(api_call: dict):
    """
    根据传入的 API 调用结构体执行对应的 API 调用。
    :param api_call: 包含 API 名称和参数信息的字典
    :return: 返回 API 调用的结果
    """
    # 获取 API 名称
    api_name = api_call.get("api_name")
    # 获取对应的函数
    api_function = api_functions.get(api_name)

    if api_function:
        # 获取参数并执行对应的 API 调用
        params = api_call.get("params", {})
        return api_function(**params)
    else:
        raise ValueError(f"未知的 API 调用：{api_name}")