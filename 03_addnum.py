import os
from openpyxl import load_workbook
from PyPDF2 import PdfFileReader, PdfFileWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from decimal import Decimal

directory = "./archiv"
excel_file = "./output.xlsx"

def add_text_to_pdf(pdf_file, text):
    # 读取原始PDF文件
    reader = PdfFileReader(pdf_file)
    writer = PdfFileWriter()
    
    # 获取第一页的内容
    page = reader.getPage(0)
    page_width = float(Decimal(page.mediaBox.getWidth()))
    page_height = float(Decimal(page.mediaBox.getHeight()))
    
    # 创建一个新的PDF页面
    c = canvas.Canvas("temp.pdf", pagesize=(page_width, page_height))
    
    # 在指定位置添加文本框
    text_x = page_width - (58 / 25.4 * 72)  # 将60mm转换为像素
    text_y = page_height - (25 / 25.4 * 72)  # 将30mm转换为像素
    text_width = 99.21
    text_height = 28.35
    c.setFont("Helvetica-Bold", 6)  # 设置字体和字号，并加粗
    c.drawString(text_x, text_y, text)  # 添加文本
    
    # 将画布内容保存到临时PDF文件
    c.save()
    
    # 读取临时PDF文件
    temp_pdf = PdfFileReader("temp.pdf")
    temp_page = temp_pdf.getPage(0)
    
    # 将文本页面与原始页面合并
    page.mergePage(temp_page)
    writer.addPage(page)
    
    # 复制剩余页面
    for i in range(1, reader.getNumPages()):
        page = reader.getPage(i)
        writer.addPage(page)
    
    # 将修改后的PDF保存到原文件
    with open(pdf_file, "wb") as output_file:
        writer.write(output_file)
    
    # 删除临时文件
    os.remove("temp.pdf")

# 加载Excel文件
workbook = load_workbook(excel_file)
sheet = workbook["Sheet2"]

# 遍历目录下的PDF文件并添加文本框
pdf_files = []
for root, dirs, files in os.walk(directory):
    for file in files:
        if file.endswith(".pdf"):
            pdf_files.append(os.path.join(root, file))

# 按顺序依次读取Excel数据并添加到PDF文件中
index = 0
for pdf_file in pdf_files:
    text = str(sheet["A"][index].value)  # 获取Excel数据
    add_text_to_pdf(pdf_file, text)
    index = (index + 1) % len(sheet["A"])
