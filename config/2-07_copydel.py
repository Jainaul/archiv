import os
import shutil

# 复制文件
source_file = "./outputchange.xlsx"
destination_folder = "./template/main"
destination_file = "./template/main/export.xlsx"

shutil.copyfile(source_file, destination_file)

# 删除源文件
os.remove(source_file)

# 删除archiv文件夹中的文件
archive_directory = "../archiv"
for root, dirs, files in os.walk(archive_directory):
    for file in files:
        if file == "outputchange.xlsx" or file == "output.xlsx":
            file_path = os.path.join(root, file)
            os.remove(file_path)

# 删除../output/doc文件夹的全部内容（包括文件和子文件夹）
output_doc_directory = "../output/doc"
if os.path.exists(output_doc_directory):
    shutil.rmtree(output_doc_directory)
