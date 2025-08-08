import re
import win32com.client
import json

def word2json(word, doc):
    # 获取文档内容和格式
    content = []

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
        para_style = para.Style.Name  # 获取段落样式

        # 判断段落是否为标题
        is_title = para_style.startswith('标题')  # 如果段落样式以 '标题' 开头，认为是标题

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
                'is_title': is_title  # 标记是否为标题
            })

    # 返回 JSON 字符串
    return json.dumps(content, ensure_ascii=False, indent=4)
