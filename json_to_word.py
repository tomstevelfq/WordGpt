import win32com.client
import json

# 启动 Word 应用程序
word = win32com.client.Dispatch("Word.Application")
word.Visible = False  # 设置为 False 以便后台运行

# 创建一个新的文档
doc = word.Documents.Add()

# 从 JSON 文件读取内容
with open('document_content.json', 'r', encoding='utf-8') as json_file:
    content = json.load(json_file)

# 恢复文档内容
for item in content:
    if item['type'] == 'paragraph':  # 如果是段落
        para = doc.Paragraphs.Add()
        para.Range.Text = item['text']
        
        # 恢复段落字体样式
        para.Range.Font.Name = item['font']
        para.Range.Font.Size = item['size']
        para.Range.Font.Bold = item['bold']
        para.Range.Font.Italic = item['italic']
        para.Range.Font.Underline = item['underline']
        
        # 恢复段落对齐和行距
        para.Alignment = item['alignment']  # 对齐方式
        para.Format.LineSpacing = item['line_spacing']  # 行距

# 保存恢复的文档
doc.SaveAs(r"C:\Users\tomst\Desktop\lunwen\wordgpt\document_recover.docx")

# 关闭文档
doc.Close()

# 退出 Word
word.Quit()

