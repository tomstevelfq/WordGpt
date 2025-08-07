import win32com.client

# 启动 Word 应用程序
def start_word():
    """ 启动 Word 应用程序并返回应用实例和文档对象 """
    global doc
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # 后台运行
    doc = word.Documents.Add()  # 创建一个新文档
    return word

# 打开现有文档
def open_document(word, file_path):
    """ 打开一个现有的 Word 文档 """
    doc = word.Documents.Open(file_path)
    return doc

# 创建新文档
def create_document(word):
    """ 创建一个新的空文档 """
    doc = word.Documents.Add()
    return doc

# 保存文档
def save_document(doc, file_path):
    """ 保存当前文档到指定路径 """
    if doc:
        doc.SaveAs(file_path)

# 关闭文档
def close_document(doc):
    """ 关闭当前文档 """
    if doc:
        doc.Close()

# 退出 Word 应用程序
def quit_word(word):
    """ 退出 Word 应用程序 """
    word.Quit()

# 添加段落并设置格式
def add_paragraph(text, style="正文", font_name="宋体", font_size=12, bold=False, italic=False, alignment=1):
    """ 添加段落并设置格式 """
    global doc
    para = doc.Content.Paragraphs.Add()  # 添加段落
    para.Range.Text = text  # 设置段落内容
    para.Style = style  # 设置段落样式
    para.Range.Font.Name = font_name  # 设置字体
    para.Range.Font.Size = font_size  # 设置字体大小
    para.Range.Font.Bold = bold  # 设置加粗
    para.Range.Font.Italic = italic  # 设置斜体
    para.Alignment = alignment  # 设置对齐方式

# 添加文档标题
def set_title(doc, title_text, style="标题 1", font_name="宋体", font_size=18):
    """ 设置文档标题 """
    title = doc.Content.Paragraphs.Add()
    title.Range.Text = title_text
    title.Style = style
    title.Range.Font.Name = font_name
    title.Range.Font.Size = font_size
    title.Range.ParagraphFormat.Alignment = 1  # 居中对齐

# 设置段落对齐方式
def set_paragraph_alignment(para, alignment):
    """ 设置段落对齐方式 """
    para.Alignment = alignment  # 0-左对齐，1-居中对齐，2-右对齐

# 设置段落字体和字号
def set_paragraph_font(para, font_name, font_size):
    """ 设置段落字体和字号 """
    para.Range.Font.Name = font_name
    para.Range.Font.Size = font_size

# 设置段落加粗
def set_paragraph_bold(para, bold):
    """ 设置段落加粗 """
    para.Range.Font.Bold = bold

# 设置段落斜体
def set_paragraph_italic(para, italic):
    """ 设置段落斜体 """
    para.Range.Font.Italic = italic

# 修改段落样式
def modify_paragraph_style(doc, paragraph_index, style="正文", font_name="宋体", font_size=12):
    """ 修改段落样式 """
    para = doc.Paragraphs(paragraph_index)
    para.Style = style
    para.Range.Font.Name = font_name
    para.Range.Font.Size = font_size

# 根据条件选择段落
def select_paragraphs(doc, filter_fn=None):
    """ 根据条件选择段落 """
    paragraphs = []
    for para in doc.Paragraphs:
        if filter_fn and filter_fn(para):
            paragraphs.append(para)
    return paragraphs

def modify_paragraph(doc, paragraph_index, new_text):
    """ 修改段落文本 """
    para = doc.Paragraphs(paragraph_index)
    para.Range.Text = new_text