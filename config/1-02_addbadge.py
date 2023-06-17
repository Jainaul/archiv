import os
from PyPDF2 import PdfFileReader, PdfFileWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from decimal import Decimal
from PIL import Image
import tempfile
import shutil

input_directory = "../archiv"
output_directory = "../output"
badge_file = "./badge.png"
badge_width = 141.73  # 指定徽章宽度为50mm
badge_height = 65.69  # 指定徽章高度为20mm

def add_badge_to_pdf(pdf_file, badge_file):
    # 读取原始PDF文件
    reader = PdfFileReader(pdf_file)
    writer = PdfFileWriter()

    # 获取第一页的内容
    page = reader.getPage(0)
    page_width = float(Decimal(page.mediaBox.getWidth()))
    page_height = float(Decimal(page.mediaBox.getHeight()))

    # 创建一个新的PDF页面
    c = canvas.Canvas("temp.pdf", pagesize=(page_width, page_height))

    # 将徽章图像保存为临时文件
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_file:
        badge_image = Image.open(badge_file)
        badge_image = badge_image.resize((int(badge_width / 25.4 * 72), int(badge_height / 25.4 * 72)))  # 将徽章大小转换为像素
        badge_image.save(tmp_file.name, "PNG")

        # 在右上角指定坐标添加徽章
        badge_x = page_width - (60 / 25.4 * 72)  # 将60mm转换为像素
        badge_y = page_height - (30 / 25.4 * 72)  # 将30mm转换为像素
        c.drawImage(tmp_file.name, badge_x, badge_y, width=badge_width, height=badge_height, mask='auto')

    # 将画布内容保存到临时PDF文件
    c.save()

    # 读取临时PDF文件
    temp_pdf = PdfFileReader("temp.pdf")
    temp_page = temp_pdf.getPage(0)

    # 将徽章页面与原始页面合并
    page.mergePage(temp_page)
    writer.addPage(page)

    # 复制剩余页面
    for i in range(1, reader.getNumPages()):
        page = reader.getPage(i)
        writer.addPage(page)

    # 构建输出文件的路径，保持相对目录结构
    relative_path = os.path.relpath(pdf_file, input_directory)
    output_file_path = os.path.join(output_directory, relative_path)

    # 确保输出目录存在
    os.makedirs(os.path.dirname(output_file_path), exist_ok=True)

    # 将修改后的PDF保存到新文件
    with open(output_file_path, "wb") as output_file:
        writer.write(output_file)

    # 删除临时文件
    os.remove("temp.pdf")
    os.remove(tmp_file.name)
    print(f"正在处理文件: {pdf_file}")  # 显示当前正在处理的文件

# 清空输出目录
shutil.rmtree(output_directory, ignore_errors=True)
os.makedirs(output_directory, exist_ok=True)

# 遍历目录下的PDF文件并添加徽章
for root, dirs, files in os.walk(input_directory):
    for file in files:
        if file.endswith(".pdf"):
            pdf_file = os.path.join(root, file)
            add_badge_to_pdf(pdf_file, badge_file)
