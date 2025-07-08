import os
import re
from openpyxl import load_workbook
from docx import Document
import pyexcel

# 配置
excel_folder = './assets'  # Excel 文件夹路径
word_template_path = './template/temp.docx'  # Word 模板路径
output_folder = './output'  # 输出 Word 文件夹

# 需要提取的字段（示例，需根据实际情况调整）
fields = {
    '工程部位': 'R10',
    '设计强度等级': 'J15',
    '墩柱': 'A21',
    '测区回弹代表值R1': 'S21',
    '测区回弹代表值R2': 'S22',
    '测区声速代表值1': 'Z21',
    '测区声速代表值2': 'Z22',
    '平测声速':'Z16',
    
    '修正为对测声速1':'AA21',
    '修正为对测声速2':'AA21',
    '测区强度代表值1':'AB21',
    '测区强度代表值2':'AB22',
    
    '构件强度推定值': 'AD21',
    '设计抗压强度等级': 'J15'
}

def trans_xls_to_xlsx(folder):
    for file in os.listdir(folder):
        if file.endswith('.xls'):
            xls_path = os.path.join(folder, file)
            xlsx_path = xls_path + 'x'
            pyexcel.save_book_as(file_name=xls_path, dest_file_name=xlsx_path)
            print(f"已转换: {file} -> {os.path.basename(xlsx_path)}")

def extract_data_from_excel(excel_path):
    wb = load_workbook(excel_path, data_only=True)
    data = {}
    for sheet in wb.worksheets:
        sheet_data = {}
        for key, cell in fields.items():
            value = sheet[cell].value
            sheet_data[key] = value
            if value == '' or value == ' ':
                raise Exception(f"数据可能出错 {sheet.title} {key}")
        data[sheet.title] = sheet_data    
    return data

def replace_placeholder_in_paragraph(paragraph, data):
    # 合并所有run的文本
    full_text = ''.join(run.text for run in paragraph.runs)
    # 替换所有占位符
    for key, value in data.items():
        placeholder = f"{{{{{key}}}}}"
        full_text = full_text.replace(placeholder, str(value))
    # 重新分配到runs
    if paragraph.runs:
        paragraph.runs[0].text = full_text
        for run in paragraph.runs[1:]:
            run.text = ''

def fill_word_template(data, output_path):
    doc = Document(word_template_path)
    # 1. 替换所有表格内的占位符
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_placeholder_in_paragraph(paragraph, data)
    # 2. 替换所有正文段落的占位符
    for paragraph in doc.paragraphs:
        replace_placeholder_in_paragraph(paragraph, data)
    doc.save(output_path)



def clean_filename(name):
    # 替换非法字符为下划线
    return re.sub(r'[\\/:#*?"<>|]', '_', name)

def main():
    # 创建输出文件夹
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    # 转换文件 xls到xlsx
    trans_xls_to_xlsx(excel_folder)
    
    # 获取文件列表
    for filename in os.listdir(excel_folder):
        
        # 只取excel文件
        if filename.endswith('.xlsx'):
            
            _filename = filename.replace(".xlsx", "")
            # 获取文件真实地址
            excel_path = os.path.join(excel_folder, filename)
            
            # 从excel中获取数据
            datas = extract_data_from_excel(excel_path)

            # 用“项目名称”作为新 Word 文件名
            for sheet_name in datas:
                safe_sheet_name = clean_filename(sheet_name)
                word_name = f"{_filename}_{safe_sheet_name}.docx"
                output_path = os.path.join(output_folder, word_name)
                fill_word_template(datas[sheet_name], output_path)
                print(f"已生成: {output_path}")

if __name__ == '__main__':
    main()