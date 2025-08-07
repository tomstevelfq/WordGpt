from llms import llm
from langchain.prompts import PromptTemplate
from langchain.chains.llm import LLMChain
from langchain.llms import BaseLLM
import json

planner_prompt="""
您将收到一个任务指令。您的任务是将这个大任务分解为一系列可以通过 API 调用逐步解决的小任务。每次规划时，您只需规划一个步骤，执行该步骤后根据执行结果继续规划后续步骤。

可用的 API 调用包括：

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
12. 设置段落斜体：`set_paragraph_italic(paragraph_index, italic)`
13. 修改段落样式：`modify_paragraph_style(paragraph_index_list, style="正文", font_name="宋体", font_size=12)`
14. 修改段落文本：`modify_paragraph(paragraph_index, new_text)`
15. 获取 Word 文档内容：`get_word_content()`

您需要根据任务指令规划每个执行步骤，并通过 API 调用来实现。每个步骤的规划应包括：

- **Plan step**: 一个详细的任务步骤。
- **API response**: 对应 API 调用的响应结果。
- **计划的每个步骤**：每次只规划一个步骤，确保每个步骤都可以用单个 API 调用来解决。

### 示例：

(1)任务指令：为用户创建一个名为“项目计划”的文档，并在文档中添加标题“项目概述”，然后保存为 `project_plan.docx`。

**规划步骤 1**：
Plan step 1: 创建一个新的空文档。
API response: 新文档已创建。

**规划步骤 2**：
Plan step 2: 设置文档标题为“项目概述”，字体为宋体，字号为 18。
API response: 文档标题已设置为“项目概述”，字体为宋体，字号为 18。

**规划步骤 3**：
Plan step 3: 保存文档为 `project_plan.docx`。
API response: 文档已保存为 `project_plan.docx`。

**最终答案**：
我已成功创建一个名为“项目计划”的文档，标题为“项目概述”，并将其保存为 `project_plan.docx`。

---


(2)任务指令：修改已有文档，将文档中的所有标题更改为斜体。

规划步骤 1：
Plan step 1: 打开现有文档 existing_document.docx。
API response: 成功打开 existing_document.docx 文档。

规划步骤 2：
Plan step 2: 获取文档内容的 JSON 格式，分析 JSON 并获取所有标题段落的索引。
API response: ```[
    {{
        "type": "paragraph",
        "text": "项目概述",
        "font": "宋体",
        "size": 18.0,
        "bold": 1,
        "italic": 0,
        "underline": 0,
        "alignment": 1,
        "line_spacing": 1.5,
        "is_title": true
    }},
    {{
        "type": "paragraph",
        "text": "工作安排",
        "font": "宋体",
        "size": 16.0,
        "bold": 0,
        "italic": 0,
        "underline": 0,
        "alignment": 0,
        "line_spacing": 1.15,
        "is_title": true
    }},
    {{
        "type": "paragraph",
        "text": "项目目标",
        "font": "宋体",
        "size": 16.0,
        "bold": 0,
        "italic": 0,
        "underline": 0,
        "alignment": 0,
        "line_spacing": 1.15,
        "is_title": false
    }}
]```

规划步骤 3：
Plan step 3: 获取到文档内容的 JSON 格式，标题段落的索引为 [1, 3, 5, 7],对每个标题段落的进行操作，将它们的样式设置为斜体。
API response: 所有标题段落（索引为 [1, 3, 5, 7]）已成功设置为斜体。

规划步骤 4：
Plan step 4: 保存修改后的文档为 existing_document_updated.docx。
API response: 文档已保存为 existing_document_updated.docx。

最终答案：
我已成功修改 existing_document.docx 中的所有标题（索引为 [1, 3, 5, 7]），将它们设置为斜体，并保存为 existing_document_updated.docx。

**注意事项**：
1. 每次规划一个步骤，不要跳过,进一步规划要等api返回才能决定。
2. 每个计划步骤应该是一个独立的 API 调用，确保任务可以按顺序执行。
3. 每次执行步骤后，根据 API 响应进行下一步规划，确保任务逐步完成。
4. 在规划步骤中，避免使用模糊的描述，要明确每个步骤的执行内容。

任务指令：{input}  
Plan step 1: {agent_scratchpad}
"""

history=[]
# 任务规划器函数：根据任务指令和历史记录生成下一个任务步骤
def construct_scratchpad():
    """
    根据历史记录构造 scratchpad
    :param history: 任务历史记录，每个记录包括执行步骤和API响应
    :return: 拼接后的 scratchpad 字符串
    """
    if len(history) == 0:
        return ""
    scratchpad = ""
    for i, (plan, execution_res) in enumerate(history):
        scratchpad += f"Plan step {i + 1}: " + plan + "\n"
        scratchpad += "API response: " + execution_res + "\n"
    return scratchpad


def generate_next_step(task_instruction):
    """
    根据任务指令和历史记录推理下一步的任务
    :param task_instruction: 当前任务指令
    :param history: 历史任务记录
    :param scenario: 当前任务的场景（例如 tmdb, spotify 等）
    :param llm: 使用的 LLM 模型
    :return: 新的任务步骤
    """
    # 根据历史记录构造 scratchpad
    scratchpad = construct_scratchpad()
    
    # 规划提示模板
    plan_prompt = PromptTemplate(
        template=planner_prompt,
        partial_variables={
            "agent_scratchpad": scratchpad,
        },
        input_variables=["input"]
    )
    
    # 使用 LLMChain 执行规划
    planner_chain = LLMChain(llm=llm, prompt=plan_prompt)
    planner_chain_output = planner_chain.run(input=task_instruction)

    # 处理返回的规划结果
    planner_chain_output = planner_chain_output.strip()
    
    # 将新的步骤添加到历史记录中
    history.append((planner_chain_output,''))
    return planner_chain_output
