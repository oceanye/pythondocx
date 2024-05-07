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

import os  # 导入操作系统功能模块，用于文件和目录操作
from docx import Document  # 从docx模块导入Document类，用于操作Word文档
from docx.shared import Pt  # 从docx.shared导入Pt，用于设置字体大小
from docx.oxml.ns import qn  # 从docx.oxml.ns导入qn，用于设置Word中的字体属性以支持中文

def sanitize_filename(filename):
    # 定义一个函数来清理文件名中的非法字符
    invalid_chars = '<>:"/\\|?*'  # 定义Windows系统中文件名不允许包含的字符
    for char in invalid_chars:
        filename = filename.replace(char, '_')  # 将非法字符替换为下划线
    return filename  # 返回清理后的文件名

def save_document(text, path, filename):
    # 定义一个函数来保存文档
    filename = sanitize_filename(filename)  # 清理文件名中的非法字符
    doc = Document()  # 创建一个新的Word文档对象
    style = doc.styles['Normal']  # 获取文档的默认样式
    font = style.font  # 获取样式中的字体设置
    font.name = '宋体'  # 设置字体为宋体
    font.size = Pt(12)  # 设置字体大小为12磅
    style.paragraph_format.first_line_indent = Pt(24)  # 设置段落首行缩进为24磅

    font.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')  # 确保中文字符能正确显示
    text = text.lstrip()  # 清除文本开始部分的空白字符
    doc.add_paragraph(text, style=style)  # 向文档中添加段落

    os.makedirs(path, exist_ok=True)  # 确保文件路径存在，如果不存在则创建
    doc.save(os.path.join(path, f"{filename}.docx"))  # 将文档保存到指定路径

def process_word_document(file_path):
    # 定义一个函数来处理Word文档
    doc = Document(file_path)  # 读取指定路径的Word文档
    current_path = []  # 初始化当前路径列表
    base_path = "output"  # 设置基本输出路径
    content_buffer = ""  # 初始化内容缓冲区
    last_heading = ""  # 初始化最后一个标题变量

    for para in doc.paragraphs:  # 遍历文档中的所有段落
        style = para.style.name  # 获取段落的样式名
        if 'Heading' in style:  # 如果样式名包含'Heading'，表示这是一个标题
            if content_buffer:  # 如果内容缓冲区不为空
                content_buffer = content_buffer.lstrip()  # 清除缓冲区开始的空白字符
                path = os.path.join(base_path, *current_path)  # 构造文件保存路径
                filename = content_buffer[:6] if len(content_buffer) >= 6 else content_buffer  # 从缓冲区内容生成文件名
                save_document(content_buffer, path, filename)  # 保存文档
                content_buffer = ""  # 清空内容缓冲区

            level = int(style.split()[-1])  # 获取标题的层级
            current_path = current_path[:level-1]  # 根据标题层级调整当前路径
            last_heading = para.text.strip().replace(':', '').replace('/', '').replace('\\', '')  # 清理标题文本
            current_path.append(last_heading)  # 将清理后的标题添加到路径中

        elif 'Body Text' in style or 'Normal' in style:  # 如果样式是正文或默认样式
            content_buffer += para.text + "\n"  # 将段落文本添加到内容缓冲区

    if content_buffer:  # 处理文档最后一部分的内容
        content_buffer = content_buffer.lstrip()  # 清除缓冲区开始的空白字符
        path = os.path.join(base_path, *current_path)  # 构造文件保存路径
        filename = content_buffer[:6] if len(content_buffer) >= 6 else content_buffer  # 从缓冲区内容生成文件名
        save_document(content_buffer, path, filename)  # 保存文档

if __name__ == "__main__":  # 当脚本作为主程序运行时
    process_word_document("通用扩初模板.docx")  # 调用处理文档的函数



























