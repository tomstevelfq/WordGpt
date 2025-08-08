import win32com.client
from word_to_json import word2json
import json

word=None
doc=None

# 打开 word 程序
def start_word(input_json):
    params = json.loads(input_json)
    place = params.get("place", None)
    print(place)
    global word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # 后台运行

# 创建新文档
def create_document(input_json):
    params = json.loads(input_json)
    global doc
    doc = word.Documents.Add()
    return "success"

# 保存文档
def save_document(input_json):
    params = json.loads(input_json)
    file_path = params.get("file_path", "")
    file_path = "C:\\Users\\tomst\\Desktop\\lunwen\\wordgpt\\" + file_path
    global doc
    if doc:
        doc.SaveAs(file_path)
    return file_path + " save success"

# 关闭文档
def close_document(input_json):
    params = json.loads(input_json)
    global doc
    if doc:
        doc.Close()

# 退出 Word 应用程序
def quit_word(input_json):
    params = json.loads(input_json)
    word.Quit()

# 添加段落并设置格式
def add_paragraph(input_json):
    params = json.loads(input_json)
    text = params.get("text", "")
    style = params.get("style", "正文")
    font_name = params.get("font_name", "宋体")
    font_size = params.get("font_size", 12)
    bold = params.get("bold", False)
    italic = params.get("italic", False)
    alignment = params.get("alignment", 1)
    
    para = doc.Content.Paragraphs.Add()  # 添加段落
    para.Range.Text = text  # 设置段落内容
    para.Style = style  # 设置段落样式
    para.Range.Font.Name = font_name  # 设置字体
    para.Range.Font.Size = font_size  # 设置字体大小
    para.Range.Font.Bold = bold  # 设置加粗
    para.Range.Font.Italic = italic  # 设置斜体
    para.Alignment = alignment  # 设置对齐方式

# 添加文档标题
def set_title(input_json):
    params = json.loads(input_json)
    title_text = params.get("title_text", "")
    style = params.get("style", "标题 1")
    font_name = params.get("font_name", "宋体")
    font_size = params.get("font_size", 18)
    
    title = doc.Content.Paragraphs.Add()
    title.Range.Text = title_text
    title.Style = style
    title.Range.Font.Name = font_name
    title.Range.Font.Size = font_size
    title.Range.ParagraphFormat.Alignment = 1  # 居中对齐
    return title_text

# 设置段落对齐方式
def set_paragraph_alignment(input_json):
    params = json.loads(input_json)
    paragraph_index = params.get("paragraph_index", 0)
    alignment = params.get("alignment", 1)
    
    para = doc.Paragraphs[paragraph_index]
    para.Alignment = alignment  # 0-左对齐，1-居中对齐，2-右对齐

# 设置多个段落的字体和字号
def set_paragraph_font(input_json):
    params = json.loads(input_json)
    paragraph_index_list = params.get("paragraph_index_list", [])  # 获取段落索引列表
    font_name = params.get("font_name", "宋体")  # 获取字体名称
    font_size = params.get("font_size", 12)  # 获取字体大小

    for paragraph_index in paragraph_index_list:
        para = doc.Paragraphs[paragraph_index]  # 请注意这里使用小括号而不是方括号
        para.Range.Font.Name = font_name  # 设置字体名称
        para.Range.Font.Size = font_size  # 设置字体大小

# 设置段落加粗
def set_paragraph_bold(input_json):
    params = json.loads(input_json)
    paragraph_index = params.get("paragraph_index", 0)
    bold = params.get("bold", False)
    
    para = doc.Paragraphs[paragraph_index]
    para.Range.Font.Bold = bold

# 设置段落斜体
def set_paragraph_italic(input_json):
    params = json.loads(input_json)
    paragraph_index_list = params.get("paragraph_index_list", [])
    italic = params.get("italic", False)
    
    for paragraph_index in paragraph_index_list:
        para = doc.Paragraphs[paragraph_index]
        para.Range.Font.Italic = italic

# # 修改段落样式
# def modify_paragraph_style(input_json):
#     params = json.loads(input_json)
#     paragraph_index = params.get("paragraph_index", 0)
#     style = params.get("style", "正文")
#     font_name = params.get("font_name", "宋体")
#     font_size = params.get("font_size", 12)
    
#     para = doc.Paragraphs[paragraph_index]
#     para.Style = style
#     para.Range.Font.Name = font_name
#     para.Range.Font.Size = font_size

# 修改多个段落的样式
def modify_paragraph_style(input_json):
    params = json.loads(input_json)
    paragraph_index_list = params.get("paragraph_index_list", [])  # 获取段落索引列表
    style = params.get("style", "正文")  # 获取段落样式
    font_name = params.get("font_name", "宋体")  # 获取字体名称
    font_size = params.get("font_size", 12)  # 获取字体大小

    for paragraph_index in paragraph_index_list:
        para = doc.Paragraphs[paragraph_index]
        para.Style = style  # 设置段落样式
        para.Range.Font.Name = font_name
        para.Range.Font.Size = font_size  # 设置字体和字号

# 修改段落文本
def modify_paragraph(input_json):
    params = json.loads(input_json)
    paragraph_index = params.get("paragraph_index", 0)
    new_text = params.get("new_text", "")
    
    para = doc.Paragraphs[paragraph_index]
    para.Range.Text = new_text

# 获取 Word 文档内容
def get_word_content(input_json):
    params = json.loads(input_json)
    return word2json(word, doc)

# 打开现有文档
def open_document(input_json):
    params = json.loads(input_json)
    file_path = params.get("file_path", "")
    print(file_path)
    if not word:
        start_word(input_json)
    global doc
    file_path = "C:\\Users\\tomst\\Desktop\\lunwen\\wordgpt\\" + file_path
    doc = word.Documents.Open(file_path)
    return get_word_content(input_json)