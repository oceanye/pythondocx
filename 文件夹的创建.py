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
md_file_path = 'Project_Scheme.md'

# 指定路径
base_path = 'Z:/BIM 内部资料/2024 03 29 项目标书自动化生成调研/初步设计说明案例'

# 创建文件夹
create_folders_from_markdown_file(md_file_path, base_path)

"""#用于检测二级标题的数量书否正确
def count_and_collect_markdown_levels(md_file_path):
    with open(md_file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    level_2_count = 0
    level_2_content = []

    for line in lines:
        line = line.strip()  # 去除空白字符（包括空行）

        if line.count('#') == 2:
            # 获取二级标题内容
            content = line.strip('# ').strip()
            # 统计二级标题个数
            level_2_count += 1
            # 保存二级标题内容
            level_2_content.append(content)

    return level_2_count, level_2_content

# Markdown 文件路径
md_file_path = 'Project_Scheme.md'

# 统计二级标题的个数及内容
level_2_count, level_2_content = count_and_collect_markdown_levels(md_file_path)

# 打印结果
print(f"二级标题个数: {level_2_count}")
print("二级标题内容:")
for content in level_2_content:
    print(content)

"""
