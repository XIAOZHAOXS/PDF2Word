"""
Author:赵文瑄
Date:2023.12.10
Power By Pycharm
"""
import os
from pdf2docx import Converter
from time import sleep

def convert_pdf_to_docx(pdf_path, docx_path):
    # 创建一个转换器实例
    cv = Converter(pdf_path)

    # 转换整个文件
    cv.convert(docx_path, start=0, end=None)

    # 关闭转换器
    cv.close()

# 目录路径
current_dir = os.getcwd()
word_dir = os.path.join(current_dir, 'Word')

# 如果Word目录不存在，则创建
if not os.path.exists(word_dir):
    os.makedirs(word_dir)

# 在当前目录中找到所有PDF文件
pdf_files = [f for f in os.listdir(current_dir) if f.endswith('.pdf')]

# 将每个PDF转换为Word文档
for pdf_file in pdf_files:
    pdf_path = os.path.join(current_dir, pdf_file)
    docx_name = os.path.splitext(pdf_file)[0] + '.docx'
    docx_path = os.path.join(word_dir, docx_name)
    convert_pdf_to_docx(pdf_path, docx_path)
    print(f"已将 '{pdf_file}' 转换为 '{docx_name}'")

print("所有PDF文件已成功转换为Word文档。")
print("\033[31m" + "Power By 小赵先生" + "\033[0m")
sleep(5)