# 导入必要的模块
from docx import Document

# 创建一个新的文档
document = Document()

# 添加文档标题
document.add_heading('Document Title', 0)

# 添加一个包含不同格式文本的段落
p = document.add_paragraph('A plain paragraph having some ')
p.add_run('bold').bold = True
p.add_run(' and some ')
p.add_run('italic.').italic = True

# 添加二级标题
document.add_heading('Heading, level 1', level=1)

# 添加一个具有特定样式的引用段落
document.add_paragraph('Intense quote', style='Intense Quote')

# 添加无序列表和有序列表项
document.add_paragraph(
    'first item in unordered list', style='List Bullet'
)
document.add_paragraph(
    'first item in ordered list', style='List Number'
)

# 注释掉添加图片的代码行
#document.add_picture('monty-truth.png', width=Inches(1.25))

# 定义数据记录，用于创建表格
records = (
    (3, '101', 'Spam'),
    (7, '422', 'Eggs'),
    (4, '631', 'Spam, spam, eggs, and spam')
)

# 添加一个表格，并设置表头
table = document.add_table(rows=1, cols=3)
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Qty'
hdr_cells[1].text = 'Id'
hdr_cells[2].text = 'Desc'

# 循环，填充表格的数据行
for qty, id, desc in records:
    row_cells = table.add_row().cells
    row_cells[0].text = str(qty)
    row_cells[1].text = id
    row_cells[2].text = desc

# 添加一个分页符
document.add_page_break()

# 保存文档
document.save('demo.docx')


from docx import Document
from docx.shared import Pt,RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn


def create_book2(name, major, school, _time):
    # 生成一个文档
    doc1 = Document()
    # 增加标题
    title = doc1.add_paragraph()
    run = title.add_run('录取通知书')
    # 设置标题的样式
    run.font.size = Pt(30)
    run.font.color.rgb = RGBColor(255, 0, 0)
    run.font.name = ''
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    # 将段落格式中的对齐方式设置为居中
    title.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # 增加内容
    doc1.add_paragraph(f'__{name}__ 同学：')
    content1 = doc1.add_paragraph(
        f'兹录取你入我校 __{major}__ 专业类学习。请凭本通知书来报道。具体时间、地点见《新生入学须知》。')
    # 设置内容样式
    content1.paragraph_format.first_line_indent = Pt(30)

    # 落款
    footer = doc1.add_paragraph(f'{school}\n')
    footer.add_run(f'{_time}')
    # 右对齐
    footer.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    # 保存文档
    doc1.save(f'07_录取通知书_{name}.docx')


if __name__ == '__main__':
    # create_book()
    create_book2('吕布', '人工智能技术', '清华大学', '二0三0年八月十号')


from win32com.client import constants, gencache


def createPdf(wordPath, pdfPath):
    """
    word转pdf
    :param wordPath: word文件路径
    :param pdfPath:  生成pdf文件路径
    """
    word = gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(wordPath, ReadOnly=1)
    doc.ExportAsFixedFormat(pdfPath,
                            constants.wdExportFormatPDF,
                            Item=constants.wdExportDocumentWithMarkup,
                            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    word.Quit(constants.wdDoNotSaveChanges)


if __name__ == "__main__":
    # 路径填写绝对路径
    createPdf(
        r'07_录取通知书_吕布.docx',
        r'10_word转换成pdf.pdf'
    )

from docx import Document


def read_word():
    # 打开文档
    doc1 = Document('人工智能.docx')
    # 读取数据-段落
    for p in doc1.paragraphs:
        print(p.text)
    # 读取表格
    for t in doc1.tables:
        for row in t.rows:
            for c in row.cells:
                print(c.text, end=' ')
            print()


if __name__ == '__main__':
    read_word()



from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table

def delete_element(element):
    """
    删除文档中的元素，无论是表格还是段落
    """
    elt = element._element
    elt.getparent().remove(elt)

def iter_block_items(parent):
    """
    生成器，用于迭代Word文档中的块级项目（段落和表格）
    """
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl

    for child in parent.element.body:
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def delete_content_between_headings(doc_path, start_heading_index, end_heading_index):
    doc = Document(doc_path)
    current_heading_index = 0
    preserve_content = False
    elements_to_delete = []

    # 遍历文档元素，标记不需要保留的元素
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph) and block.style.name.startswith('Heading'):
            current_heading_index += 1
            if current_heading_index == start_heading_index:
                preserve_content = True  # 开始保留内容
            elif current_heading_index == end_heading_index:
                preserve_content = False  # 结束保留内容

        if not preserve_content:
            elements_to_delete.append(block)

    # 删除标记的元素
    for element in elements_to_delete:
        delete_element(element)

    doc.save('modified_document.docx')

delete_content_between_headings('人工智能 - 副本.docx', 1, 2)





