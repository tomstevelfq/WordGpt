import win32com.client

# 启动 Word 应用程序
word = win32com.client.Dispatch("Word.Application")
word.Visible = True  # 显示 Word 窗口

# 创建一个新的文档
doc = word.Documents.Add()

# 设置文档的标题（大标题）
title = doc.Content.Paragraphs.Add()  # 添加标题段落
title.Range.Text = "基于知识图谱和动态任务规划框架的电网故障应对机制"  # 设置标题文本
title.Style = "标题 1"  # 设置为 Word 的标题 1 样式
title.Range.Font.Size = 18  # 设置标题字体大小
title.Range.Font.Name = "宋体"  # 设置标题字体
title.Range.ParagraphFormat.Alignment = 1  # 设置标题居中对齐

# 设置小标题
subtitle = doc.Content.Paragraphs.Add()  # 添加小标题段落
subtitle.Range.Text = "步骤 1：数据准备与API设计"  # 设置小标题文本
subtitle.Style = "标题 2"  # 设置为 Word 的标题 2 样式
subtitle.Range.Font.Size = 14  # 设置小标题字体大小
subtitle.Range.Font.Name = "宋体"  # 设置小标题字体
subtitle.Range.Font.Bold = True  # 设置小标题加粗
subtitle.Range.ParagraphFormat.Alignment = 1  # 设置小标题居中对齐

# 添加段落内容
paragraph1 = doc.Content.Paragraphs.Add()  # 添加段落 1
paragraph1.Range.Text = "第一步：获取电网故障预测数据集。"  # 设置段落文本
paragraph1.Style = "正文"  # 设置为 Word 的正文样式
paragraph1.Range.Font.Size = 12  # 设置段落字体大小
paragraph1.Range.Font.Name = "宋体"  # 设置段落字体
paragraph1.Range.Font.Bold = True  # 设置段落加粗

# 添加另一段内容
paragraph2 = doc.Content.Paragraphs.Add()  # 添加段落 2
paragraph2.Range.Text = "收集包含电网故障与故障预案的多维度数据集。数据集可能包括电网设备状态、负荷、温度、湿度、电压、电流，故障应对措施等信息，以及历史故障记录（故障发生时间、地点、类型等）。"  # 设置段落文本
paragraph2.Style = "正文"  # 设置为 Word 的正文样式
paragraph2.Range.Font.Size = 12  # 设置段落字体大小
paragraph2.Range.Font.Name = "微软雅黑"  # 设置段落字体
paragraph2.Range.Font.Italic = True  # 设置字体为斜体
paragraph1.Range.Font.Bold = False  # 设置段落加粗

# 添加一个加粗的段落
bold_paragraph = doc.Content.Paragraphs.Add()  # 添加加粗段落
bold_paragraph.Range.Text = "这是加粗的段落。"  # 设置段落文本
bold_paragraph.Style = "正文"  # 设置为 Word 的正文样式
bold_paragraph.Range.Font.Bold = True  # 设置加粗

# 保存文档
doc.SaveAs(r"C:\Users\tomst\Desktop\lunwen\wordgpt\wordtest_with_details.docx")

# 退出 Word 应用程序
word.Quit()
