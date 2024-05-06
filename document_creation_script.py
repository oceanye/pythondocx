import glob
import os


def create_folders_from_markdown_file(md_file_path, base_path):
    with open(md_file_path, 'r', encoding='utf-8') as file:
        markdown_text = file.read()

    lines = markdown_text.split('\n')
    current_path = base_path
    level_stack = []  # 保存各级标题对应的文件夹路径

    for line in lines:
        line = line.strip()  # 去除空白字符（包括空行）
        if not line:
            continue  # 如果是空行，则跳过

        level = len(line) - len(line.lstrip('#'))
        folder_name = line.lstrip('# ').strip()

        # 当前标题的文件夹路径
        current_folder_path = os.path.join(current_path, folder_name)

        # 将当前标题的文件夹路径放入正确的层级位置
        if level == len(level_stack):
            # 与栈顶层级相同，说明在同一层级
            pass
        elif level > len(level_stack):
            # 比栈顶层级高，进入下一级
            level_stack.append(current_path)
        else:
            # 比栈顶层级低，回到对应层级
            level_stack = level_stack[:level]
        # 更新当前路径
        current_path = os.path.join(*level_stack, folder_name)
        # 创建文件夹
        os.makedirs(current_path, exist_ok=True)


# Markdown 文件路径
md_file_path = '11.md'

# 指定路径
base_path = 'F:/test'

# 创建文件夹
create_folders_from_markdown_file(md_file_path, base_path)
# 获取文件夹路径下的所有文件名组成的列表
j=0
#path = 'F:/test'
# for root, dirs, files in os.walk('F:/test'):
#     for i in dirs:
#         print(root+i)
#         j+=1
# print(f'文件个数：{j}')


def build_keyword_to_folder_mapping(base_path):
    keyword_to_folder = {}
    # 遍历base_path下的所有文件夹
    for root, dirs, files in os.walk(base_path, topdown=True):
        for name in dirs:
            # 构建完整的路径
            folder_path = os.path.join(root, name)
            # 将文件夹名称作为关键字，映射到其完整路径
            keyword_to_folder[root] = folder_path
    return keyword_to_folder

# 构建映射字典
keyword_to_folder = build_keyword_to_folder_mapping(base_path)

# 打印出构建的字典以查看结果
print("Keyword to Folder Mapping:")
for keyword, folder_path in keyword_to_folder.items():
    print(f"Keyword: {keyword}, Folder: {folder_path}")

# 现在可以使用这个字典来处理文档内容
# 实现按标题建立文件夹，且拆分文档段落


import os
from docx import Document
from docx.shared import Pt  # 导入Pt用于定义字号
from docx.oxml.ns import qn  # 导入qn用于处理中文格式问题

def save_document(text, path, filename):
    # 创建一个新的 Word 文档并添加正文
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = '宋体'
    font.size = Pt(12)
    style.paragraph_format.first_line_indent = Pt(24)  # 首行缩进2个字符（大约是24磅）

    # 设置字体支持中文（在某些Word版本中需要）
    font.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # 清除文本中的前导空格后再添加到文档
    text = text.lstrip()
    doc.add_paragraph(text, style=style)

    # 确保目录存在
    os.makedirs(path, exist_ok=True)
    # 保存文档
    doc.save(os.path.join(path, f"{filename}.docx"))

def process_word_document(file_path):
    doc = Document(file_path)
    current_path = []
    base_path = "output"  # 根输出文件夹
    content_buffer = ""  # 缓存标题下的正文内容
    last_heading = ""    # 保存上一个标题，用于文件命名

    for para in doc.paragraphs:
        style = para.style.name
        if 'Heading' in style:
            if content_buffer:  # 当存在缓存的正文时，保存到文件
                content_buffer = content_buffer.lstrip()  # 去掉缓存内容前的空格
                path = os.path.join(base_path, *current_path)
                # 使用正文内容的前6个字命名文件，确保正文长度足够
                filename = content_buffer[:6] if len(content_buffer) >= 6 else content_buffer
                save_document(content_buffer, path, filename)
                content_buffer = ""  # 重置正文缓存

            # 处理标题，更新路径和标题名称
            level = int(style.split()[-1])
            current_path = current_path[:level-1]
            last_heading = para.text.strip().replace(':', '').replace('/', '').replace('\\', '')
            current_path.append(last_heading)

        elif 'Body Text' in style or 'Normal' in style:
            # 添加正文到缓存
            content_buffer += para.text + "\n"

    # 处理最后一个标题下的内容
    if content_buffer:
        content_buffer = content_buffer.lstrip()  # 去掉缓存内容前的空格
        path = os.path.join(base_path, *current_path)
        filename = content_buffer[:6] if len(content_buffer) >= 6 else content_buffer
        save_document(content_buffer, path, filename)

if __name__ == "__main__":
    process_word_document("人工智能.docx")























