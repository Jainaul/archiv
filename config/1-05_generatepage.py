import os
from PyPDF2 import PdfFileReader, PdfFileWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

directory = "../output"

def add_page_numbers(directory):
    for root, dirs, files in os.walk(directory):
        for dir in dirs:
            subdir = os.path.join(root, dir)
            add_page_numbers_to_folder(subdir)

def add_page_numbers_to_folder(folder):
    pdf_files = get_pdf_files_in_folder(folder)
    start_page = 1
    
    for pdf_file in pdf_files:
        total_pages = get_total_pages(pdf_file)
        add_page_numbers_to_pdf(pdf_file, total_pages, start_page)
        start_page += total_pages

def get_pdf_files_in_folder(folder):
    pdf_files = []
    for file in os.listdir(folder):
        if file.endswith(".pdf"):
            pdf_file = os.path.join(folder, file)
            pdf_files.append(pdf_file)
    return pdf_files

def get_total_pages(pdf_file):
    reader = PdfFileReader(pdf_file)
    total_pages = reader.getNumPages()
    return total_pages

def add_page_numbers_to_pdf(pdf_file, total_pages, start_page):
    output_file = "temp.pdf"
    c = canvas.Canvas(output_file, pagesize=letter)
    
    # 设置字体和字号
    c.setFont("Helvetica", 12)
    
    # 获取每页的页码并添加到PDF文件中
    for page_num in range(1, total_pages + 1):
        c.setFont("Helvetica", 12)
        c.drawString(400, 10, str(start_page + page_num - 1))  # 在指定位置添加页码
        
        c.showPage()  # 创建新页面
    
    c.save()
    
    # 合并页码到原始PDF文件
    merge_pdf_files(pdf_file, output_file)
    
    # 删除临时文件
    os.remove(output_file)
    print(f"正在处理文件: {pdf_file}")  # 显示当前正在处理的文件

def merge_pdf_files(original_pdf_file, page_number_pdf_file):
    original_pdf = PdfFileReader(original_pdf_file)
    page_number_pdf = PdfFileReader(page_number_pdf_file)
    
    # 创建新的PDF文件写入器
    writer = PdfFileWriter()
    
    # 将每个页码页面添加到原始PDF文件中的每一页
    for i in range(original_pdf.getNumPages()):
        page = original_pdf.getPage(i)
        page_number_page = page_number_pdf.getPage(i % page_number_pdf.getNumPages())
        page.mergePage(page_number_page)
        writer.addPage(page)
    
    # 将修改后的PDF保存到原文件
    with open(original_pdf_file, "wb") as output_file:
        writer.write(output_file)

# 添加页码到PDF文件夹
add_page_numbers(directory)
