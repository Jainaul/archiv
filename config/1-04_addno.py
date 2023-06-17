import os
from PyPDF2 import PdfFileReader, PdfFileWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from decimal import Decimal

directory = "../output"

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
    text_x = page_width - (19 / 25.4 * 72)  # 将30mm转换为像素
    text_y = page_height - (25 / 25.4 * 72)  # 将20mm转换为像素
    text_width = 99.21
    text_height = 28.35
    c.setFont("Helvetica", 12)  # 设置字体和字号
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
    print(f"正在处理文件: {pdf_file}")  # 显示当前正在处理的文件

# 遍历目录下的PDF文件并添加文本框
for root, dirs, files in os.walk(directory):
    file_counter = 1
    for file in files:
        if file.endswith(".pdf"):
            pdf_file = os.path.join(root, file)
            text = str(file_counter)  # 获取文本数据，这里使用文件计数作为文本
            add_text_to_pdf(pdf_file, text)
            file_counter += 1
