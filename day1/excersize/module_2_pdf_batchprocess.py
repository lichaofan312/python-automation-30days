# 场景1： Excel 转 PDF（无需手动打开 Excel 导出）
# 很多时候需要将 Excel 报表转成 PDF 发给客户，用代码实现批量转换

import os
import re

from PyPDF2 import PdfReader, PdfWriter
from win32com.client import DispatchEx


def xlsx2pdf(filename):
    try:
        xlApp = DispatchEx("Excel.Application")
        # 后台运行, 不显示, 不警告
        xlApp.Visible = False
        xlApp.DisplayAlerts = 0
        books = xlApp.Workbooks.Open(filename, False)
        # 第一个参数0表示转换pdf
        books.ExportAsFixedFormat(0, re.subn('.xlsx', '.pdf', filename)[0])
        books.Close(False)
        print('保存 PDF 文件：', re.subn('.xlsx', '.pdf', filename)[0])
    except:
        print('except')
        input('转换出错了，按任意键退出')
    finally:
        print('finally')
        xlApp.Quit()


def merge_pdfs(input_dir, output_path):
    pdf_files = [f for f in os.listdir(input_dir) if f.endswith('.pdf')]
    pdf_writer = PdfWriter()
    print(f'input_dir:{input_dir}')
    print(f'pdf_files:{pdf_files}')
    for pdf_file in sorted(pdf_files):
        pdf_path = os.path.join(input_dir, pdf_file)
        print(f'pdf_path: {pdf_path}')
        pdf_reader = PdfReader(pdf_path)
        print(f'pdf_reader.pages:{len(pdf_reader.pages)}')
        for page_num, page in enumerate(pdf_reader.pages):
            print(f'page_num:{page_num}')
            # page = pdf_reader.pages(page_num)
            pdf_writer.add_page(page)
        print()

    with open(output_path, 'wb') as out:
        pdf_writer.write(out)


def gen_filepath():
    filepath = input('输入你的文件路径：')
    if not filepath:
        print('1')
        filepath = r'D:\Users\Administrator\Desktop\pythonLib\虚拟环境\project4\excersize'
    else:
        print('2')
    return filepath


# D:\Users\Administrator\Desktop\pythonLib\虚拟环境\project4\excersize
if __name__ == '__main__':
    print(f'输入 0 调用 xlsx2pdf 用于 excel 转 pdf')
    print(f'输入 1 调用 merge_pdfs 用于 合并pdf')

    num = int(input('输入你的功能编号: '))
    if num == 0:
        filepath = gen_filepath()
        for dirs, subdirs, files in os.walk(filepath):
            print(f'dirs:{dirs}')
            print(f'subdirs:{subdirs}')
            print(f'files:{files}')
            for name in files:
                print(f'name:{name}')
                if re.search('.xlsx', name):
                    xlsx2pdf(filepath + '\\' + name)
        input('转换成功，按任意键推出')
    else:
        input_directory = gen_filepath()
        output_pdf = input_directory + '/merged.pdf'
        merge_pdfs(input_directory, output_pdf)
