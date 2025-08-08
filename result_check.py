from llms import llm
from langchain.prompts import PromptTemplate
import json
from langchain.chains import LLMChain

result_check_prompt="""
任务要求：{task_instruction}

文档内容（JSON格式）如下：
{word_json}

### 示例判断：
假设任务要求为：修改文档中的所有标题为斜体，字体为宋体，字号为 22。

假设文档内容为：
[
    {{
        "type": "paragraph",
        "text": "基于知识图谱和动态任务规划框架的电网故障应对机制",
        "font": "宋体",
        "size": 22.0,
        "bold": -1,
        "italic": 1,
        "underline": 0,
        "alignment": 1,
        "line_spacing": 28.8,
        "is_title": true
    }},
    {{
        "type": "paragraph",
        "text": "步骤 1：数据准备与API设计",
        "font": "宋体",
        "size": 22.0,
        "bold": -1,
        "italic": 1,
        "underline": 0,
        "alignment": 1,
        "line_spacing": 20.6,
        "is_title": true
    }}
]

输出：
{{
    "result": "yes",
    "reason": "根据上述文档内容和任务要求，文档满足任务要求，因为所有标题的字体为“宋体”，字号为 22，并且斜体已设置。"
}}

请根据任务要求和文档内容，判断文档是否已完成任务。如果文档满足任务要求，返回 "yes"，否则返回 "no"。

注意：输出只能有json格式的结构体,不能出现其它内容

"""

# 使用 PromptTemplate 来格式化提示词
custom_prompt = PromptTemplate(
    template=result_check_prompt,
    input_variables=["task_instruction", "word_json"]
)

# 交互函数
def check_task_completion(task_instruction, word_json):
    # 创建 LLMChain 实例
    planner_chain = LLMChain(llm=llm, prompt=custom_prompt)

    # 运行链并获取结果
    planner_chain_output = planner_chain.run(task_instruction=task_instruction,word_json=word_json)

    # 获取返回的结果
    return planner_chain_output

word_json="""
[
    {
        "type": "paragraph",
        "text": "基于知识图谱和动态任务规划框架的电网故障应对机制",
        "font": "宋体",
        "size": 22.0,
        "bold": -1,
        "italic": 0,
        "underline": 0,
        "alignment": 1,
        "line_spacing": 28.80000114440918,
        "is_title": true
    },
    {
        "type": "paragraph",
        "text": "步骤 1：数据准备与API设计",
        "font": "",
        "size": 16.0,
        "bold": -1,
        "italic": 0,
        "underline": 0,
        "alignment": 1,
        "line_spacing": 20.649999618530273,
        "is_title": true
    },
    {
        "type": "paragraph",
        "text": "第一步：获取电网故障预测数据集。",
        "font": "宋体",
        "size": 12.0,
        "bold": -1,
        "italic": 0,
        "underline": 0,
        "alignment": 3,
        "line_spacing": 12.0,
        "is_title": false
    },
    {
        "type": "paragraph",
        "text": "收集包含电网故障与故障预案的多维度数据集。数据集可能包括电网设备状态、负荷、温度、湿度、电压、电流，故障应对措施等信息，以及历史故障记录（故障发生时间、地点、类型等）。",
        "font": "微软雅黑",
        "size": 12.0,
        "bold": 0,
        "italic": -1,
        "underline": 0,
        "alignment": 3,
        "line_spacing": 12.0,
        "is_title": false
    },
    {
        "type": "paragraph",
        "text": "这是加粗的段落。",
        "font": "微软雅黑",
        "size": 12.0,
        "bold": -1,
        "italic": -1,
        "underline": 0,
        "alignment": 3,
        "line_spacing": 12.0,
        "is_title": false
    }
]
"""

def extract_result(json_string):
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
        return data.get("result", "no")
    except json.JSONDecodeError:
        # 如果 JSON 解析失败，返回空列表
        print("Error decoding JSON.")
        return "no"

print(extract_result(check_task_completion("修改文档中的所有标题为微软雅黑,20号",word_json)))

