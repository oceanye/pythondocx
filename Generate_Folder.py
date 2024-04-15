import os
from docx import Document

def create_folder_structure(doc, parent_dir=''):
    # 遍历文档中的所有段落
    for paragraph in doc.paragraphs:
        print(paragraph.style.name)
        # 检查段落样式是否为标题样式
        if paragraph.style.name.startswith('Heading'):
            # 从样式名称中提取标题级别
            # level 是pargraph的最后一位数字
            level = int(paragraph.style.name[-1])
            # 获取标题文本
            title_text = paragraph.text.strip().replace('\n', ' ').replace('\r', '')
            # 构建新文件夹的路径
            new_dir = os.path.join(parent_dir, title_text)
            # 创建新文件夹
            os.makedirs(new_dir, exist_ok=True)

# 读取Word文档
doc = Document('Design Specifications.docx')

# 从文档的一级标题和二级标题开始创建文件夹结构
create_folder_structure(doc)