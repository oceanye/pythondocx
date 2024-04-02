from docx import Document

# 定义要读取和写入的文件路径
source_path = '/素材/PC外墙/PC外墙介绍.docx'
target_path = 'styled_document.docx'

#当前路径下新建source_pat

# 读取源文档
source_doc = Document(source_path)
source_text = ''
for paragraph in source_doc.paragraphs:
    source_text += paragraph.text + '\n'  # 将读取的每一段文字添加到source_text字符串中

# 打开目标文档
target_doc = Document(target_path)

# 定义或获取样式3
# 假设你已经有了样式3定义在target_doc中，否则你需要像之前那样创建样式
style3 = target_doc.styles['样式3']

# 将读取的文字添加到目标文档，并应用样式3
paragraph = target_doc.add_paragraph(source_text)
paragraph.style = style3

# 保存更改
target_doc.save(target_path)
