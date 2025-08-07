import win32com.client
import json

# 启动 Word 应用程序
word = win32com.client.Dispatch("Word.Application")
word.Visible = False  # 设置为 False 以便后台运行

# 打开现有文档
doc = word.Documents.Open(r"C:\Users\tomst\Desktop\lunwen\wordgpt\document_example.docx")

# 获取文档内容和格式
content = []

# 用于跟踪已处理的表格（确保不重复处理）
processed_tables = set()

# 遍历所有段落
for para in doc.Paragraphs:
    para_text = para.Range.Text.strip()  # 获取段落文本
    para_font = para.Range.Font.Name  # 获取字体
    para_size = para.Range.Font.Size  # 获取字体大小
    para_bold = para.Range.Font.Bold  # 获取字体是否加粗
    para_italic = para.Range.Font.Italic  # 获取字体是否斜体
    para_underline = para.Range.Font.Underline  # 获取字体是否下划线
    para_alignment = para.Alignment  # 获取段落对齐方式
    para_line_spacing = para.Format.LineSpacing  # 获取行距
    para_style = para.Style.NameLocal  # 获取段落的样式

    # 判断段落是否为标题
    is_title = para_style.lower().startswith("标题")  # 如果样式名称是 "Heading 1", "Heading 2", 等，认为是标题

    # 如果段落文本非空且不是包含控制字符的文本
    if para_text:  # 忽略无效的控制字符
        content.append({
            'type': 'paragraph',
            'text': para_text,
            'font': para_font,
            'size': para_size,
            'bold': para_bold,
            'italic': para_italic,
            'underline': para_underline,
            'alignment': para_alignment,
            'line_spacing': para_line_spacing,
            'is_title': is_title  # 添加是否是标题的属性
        })
        
# 将内容保存为 JSON 文件
with open('document_content.json', 'w', encoding='utf-8') as json_file:
    json.dump(content, json_file, ensure_ascii=False, indent=4)

# 关闭文档
doc.Close()