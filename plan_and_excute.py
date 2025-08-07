from planner import *
from executor import *

task_instruct="打开document_example.docx，将所有标题改成斜体，所有段落改成微软雅黑"
while(True):
    planner_out_put=generate_next_step(task_instruct)
    api_call_json=execute_task(planner_out_put)
    api_call_result=execute_api_call(api_call_json)
    if "结束" in api_call_json or "结束" in api_call_result:
        break

