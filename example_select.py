from llms import llm
from langchain.prompts import PromptTemplate
import json
from langchain.chains import LLMChain

task_plans = {
    "修改已有文档，将文档中的所有标题更改为斜体": {
        "steps": [
            {
                "Thought": "我需要逐步执行任务：打开文档、获取内容、修改标题样式和保存文档。",
                "Action": "open_document()",
                "Action Input": "existing_document.docx",
                "Observation": "成功打开 existing_document.docx 文档。"
            },
            {
                "Thought": "文档已打开，现在我需要获取文档内容。",
                "Action": "get_word_content()",
                "Action Input": "无",
                "Observation": "成功获取文档内容的 JSON 格式。"
            },
            {
                "Thought": "我获取到文档内容，现在我需要操作所有标题段落的样式。",
                "Action": "modify_paragraph_style()",
                "Action Input": "所有标题段落设置为斜体",
                "Observation": "所有标题段落（索引为 [1, 3, 5, 7]）已成功设置为斜体。"
            },
            {
                "Thought": "所有标题已修改，现在我需要保存修改后的文档。",
                "Action": "save_document()",
                "Action Input": "existing_document_updated.docx save success",
                "Observation": "文档已保存为 existing_document_updated.docx。"
            },
            {
                "Final Answer": "我已成功修改 existing_document.docx 中的所有标题，将它们设置为斜体，并保存为 existing_document_updated.docx。"
            }
        ]
    },
    "为文档添加标题并保存": {
        "steps": [
            {
                "Thought": "我需要逐步执行任务：打开文档、设置标题、保存文档。",
                "Action": "open_document()",
                "Action Input": "existing_document.docx",
                "Observation": "成功打开 existing_document.docx 文档。"
            },
            {
                "Thought": "文档已打开，现在我需要设置标题。",
                "Action": "set_title()",
                "Action Input": "{\"title_text\": \"项目概述\", \"style\": \"标题 1\", \"font_name\": \"宋体\", \"font_size\": 18}",
                "Observation": "标题已成功设置为“项目概述”，样式为“标题 1”。"
            },
            {
                "Thought": "标题已设置，现在我需要保存修改后的文档。",
                "Action": "save_document()",
                "Action Input": "existing_document_with_title.docx",
                "Observation": "文档已保存为 existing_document_with_title.docx。"
            },
            {
                "Final Answer": "我已成功为 existing_document.docx 添加了标题“项目概述”，并将其保存为 existing_document_with_title.docx。"
            }
        ]
    },
    "修改文档中的所有正文段落的字体为“Arial”，字号为 14": {
        "steps": [
            {
                "Thought": "我需要逐步执行任务：打开文档、获取内容、修改段落样式和保存文档。",
                "Action": "open_document()",
                "Action Input": "existing_document.docx",
                "Observation": "成功打开 existing_document.docx 文档。"
            },
            {
                "Thought": "文档已打开，现在我需要获取文档内容。",
                "Action": "get_word_content()",
                "Action Input": "无",
                "Observation": "成功获取文档内容的 JSON 格式。"
            },
            {
                "Thought": "我获取到文档内容，现在我需要修改所有正文段落的字体和字号。",
                "Action": "modify_paragraph_style()",
                "Action Input": "{\"paragraph_index_list\": [2, 4, 6], \"font_name\": \"Arial\", \"font_size\": 14}",
                "Observation": "正文段落（索引为 [2, 4, 6]）已成功设置为 Arial 字体，字号为 14。"
            },
            {
                "Thought": "所有正文段落已修改，现在我需要保存修改后的文档。",
                "Action": "save_document()",
                "Action Input": "existing_document_modified.docx",
                "Observation": "文档已保存为 existing_document_modified.docx。"
            },
            {
                "Final Answer": "我已成功修改 existing_document.docx 中的所有正文段落的字体为“Arial”，字号为 14，并保存为 existing_document_modified.docx。"
            }
        ]
    },
    "删除文档中的所有空段落": {
        "steps": [
            {
                "Thought": "我需要逐步执行任务：打开文档、获取内容、删除空段落和保存文档。",
                "Action": "open_document()",
                "Action Input": "existing_document.docx",
                "Observation": "成功打开 existing_document.docx 文档。"
            },
            {
                "Thought": "文档已打开，现在我需要获取文档内容。",
                "Action": "get_word_content()",
                "Action Input": "无",
                "Observation": "成功获取文档内容的 JSON 格式。"
            },
            {
                "Thought": "我获取到文档内容，现在我需要查找并删除空段落。",
                "Action": "modify_paragraph()",
                "Action Input": "{\"paragraph_index_list\": [5, 7, 9], \"new_text\": \"\"}",
                "Observation": "空段落（索引为 [5, 7, 9]）已成功删除。"
            },
            {
                "Thought": "空段落已删除，现在我需要保存修改后的文档。",
                "Action": "save_document()",
                "Action Input": "existing_document_no_empty_paragraphs.docx",
                "Observation": "文档已保存为 existing_document_no_empty_paragraphs.docx。"
            },
            {
                "Final Answer": "我已成功删除 existing_document.docx 中的所有空段落，并保存为 existing_document_no_empty_paragraphs.docx。"
            }
        ]
    },
    "将文档中的所有段落文本设置为加粗": {
        "steps": [
            {
                "Thought": "我需要逐步执行任务：打开文档、获取内容、设置加粗并保存文档。",
                "Action": "open_document()",
                "Action Input": "existing_document.docx",
                "Observation": "成功打开 existing_document.docx 文档。"
            },
            {
                "Thought": "文档已打开，现在我需要获取文档内容。",
                "Action": "get_word_content()",
                "Action Input": "无",
                "Observation": "成功获取文档内容的 JSON 格式。"
            },
            {
                "Thought": "我获取到文档内容，现在我需要设置所有段落文本为加粗。",
                "Action": "set_paragraph_bold()",
                "Action Input": "{\"paragraph_index_list\": [1, 2, 3, 4], \"bold\": true}",
                "Observation": "所有段落（索引为 [1, 2, 3, 4]）已成功设置为加粗。"
            },
            {
                "Thought": "所有段落文本已加粗，现在我需要保存修改后的文档。",
                "Action": "save_document()",
                "Action Input": "existing_document_bold.docx",
                "Observation": "文档已保存为 existing_document_bold.docx。"
            },
            {
                "Final Answer": "我已成功将 existing_document.docx 中的所有段落文本设置为加粗，并保存为 existing_document_bold.docx。"
            }
        ]
    },
    "修改文档标题的字体和大小": {
        "steps": [
            {
                "Thought": "我需要逐步执行任务：打开文档、获取内容、修改标题样式并保存文档。",
                "Action": "open_document()",
                "Action Input": "existing_document.docx",
                "Observation": "成功打开 existing_document.docx 文档。"
            },
            {
                "Thought": "文档已打开，现在我需要获取文档内容。",
                "Action": "get_word_content()",
                "Action Input": "无",
                "Observation": "成功获取文档内容的 JSON 格式。"
            },
            {
                "Thought": "我获取到文档内容，现在我需要修改所有标题段落的字体和大小。",
                "Action": "modify_paragraph_style()",
                "Action Input": "{\"paragraph_index_list\": [1, 3, 5], \"font_name\": \"Times New Roman\", \"font_size\": 16}",
                "Observation": "所有标题段落（索引为 [1, 3, 5]）已成功修改为“Times New Roman”字体，字号为 16。"
            },
            {
                "Thought": "所有标题已修改，现在我需要保存修改后的文档。",
                "Action": "save_document()",
                "Action Input": "existing_document_with_updated_titles.docx",
                "Observation": "文档已保存为 existing_document_with_updated_titles.docx。"
            },
            {
                "Final Answer": "我已成功将 existing_document.docx 中的所有标题段落的字体修改为“Times New Roman”，字号为 16，并保存为 existing_document_with_updated_titles.docx。"
            }
        ]
    }
}

task_keys = list(task_plans.keys())

task_select_prompt="""
你有一个任务列表 `task_keys`，其中包含多个任务要求（键）。现在，请根据输入的任务指令（`input`），从 `task_keys` 中选择出最相关的三个任务要求，并将它们以 JSON 格式返回。返回的 JSON 格式应包含与输入任务相关的任务要求的三个键，按照相关性排序。

任务列表 `task_keys` 如下：
{task_keys}

任务指令 `input` 为：{input}

请输出与 `input` 最相关的三个任务要求的键，并以 JSON 格式返回，格式如下：
{{
    "related_tasks": [
        "任务要求1",
        "任务要求2",
        "任务要求3"
    ]
}}

注意：你只能输出json，不能有其它多余内容
"""

def extract_related_tasks(json_string):
    # 如果字符串以 ```json 开头，则移除该部分
    if json_string.startswith("```json"):
        json_string = json_string[len("```json"):].strip()

    # 如果字符串以 ``` 结尾，也要去掉它
    if json_string.endswith("```"):
        json_string = json_string[:-3].strip()

    try:
        # 解析 JSON 字符串
        data = json.loads(json_string)
        
        # 返回 related_tasks 列表
        return data.get("related_tasks", [])
    except json.JSONDecodeError:
        # 如果 JSON 解析失败，返回空列表
        print("Error decoding JSON.")
        return []

def query_task_plan(input_task, task_keys):
    # 使用 PromptTemplate 来格式化提示词
    custom_prompt = PromptTemplate(
        template=task_select_prompt,
        input_variables=["input", "task_keys"]
    )

    # 创建 LLMChain 实例
    planner_chain = LLMChain(llm=llm, prompt=custom_prompt)

    # 运行链并获取结果
    planner_chain_output = planner_chain.run(input=input_task,task_keys=json.dumps(task_keys, ensure_ascii=False, indent=4))

    return extract_related_tasks(planner_chain_output)

# 函数实现
def select_tasks(task_list):
    # 从 task_plans 中选择与 task_list 中任务要求相关的项
    selected_tasks = {task: task_plans[task] for task in task_list if task in task_plans}
    
    # 将选中的任务转换为 JSON 字符串并返回
    return json.dumps(selected_tasks, ensure_ascii=False, indent=4)

def select_tasks_by_task(task):
    task_list = query_task_plan(input_task, task_keys)
    return select_tasks(task_list)

input_task = "修改文档中的标题样式"

# 打印返回的任务规划
print(select_tasks_by_task(input_task))
