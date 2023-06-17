import os
import shutil

# 源文件目录
source_directory1 = "./template/main/cover"
source_directory2 = "./template/main/list15"
source_directory3 = "./template/main/side"
source_directory4 = "./template/mini/perpar"
source_directory5 = "./template/mini/sign"
source_directory6 = "./template/minor/catalog15"

# 目标目录
target_directory1 = "../output/doc/封面"
target_directory2 = "../output/doc/移交清单"
target_directory3 = "../output/doc/侧边贴条"
target_directory4 = "../output/doc/备考表"
target_directory5 = "../output/doc/移交签字"
target_directory6 = "../output/doc/目录"

# 要删除文件的目录路径
directory_path = "./template"

# 创建目标目录
os.makedirs(target_directory1, exist_ok=True)
os.makedirs(target_directory2, exist_ok=True)
os.makedirs(target_directory3, exist_ok=True)
os.makedirs(target_directory4, exist_ok=True)
os.makedirs(target_directory5, exist_ok=True)
os.makedirs(target_directory6, exist_ok=True)

# 移动 ./template/main/cover 目录下的文件
files1 = os.listdir(source_directory1)
for file1 in files1:
    if file1 != "temp.docx":
        source_file1 = os.path.join(source_directory1, file1)
        target_file1 = os.path.join(target_directory1, file1)
        shutil.move(source_file1, target_file1)

# 移动 ./template/main/list15 目录下的文件
files2 = os.listdir(source_directory2)
for file2 in files2:
    if file2 != "temp.docx":
        source_file2 = os.path.join(source_directory2, file2)
        target_file2 = os.path.join(target_directory2, file2)
        shutil.move(source_file2, target_file2)

# 移动 ./template/main/side 目录下的文件
files3 = os.listdir(source_directory3)
for file3 in files3:
    if file3 != "temp.docx":
        source_file3 = os.path.join(source_directory3, file3)
        target_file3 = os.path.join(target_directory3, file3)
        shutil.move(source_file3, target_file3)

# 移动 ./template/mini/perpar 目录下的文件
files4 = os.listdir(source_directory4)
for file4 in files4:
    if file4 != "temp.docx":
        source_file4 = os.path.join(source_directory4, file4)
        target_file4 = os.path.join(target_directory4, file4)
        shutil.move(source_file4, target_file4)

# 移动 ./template/mini/sign 目录下的文件
files5 = os.listdir(source_directory5)
for i, file5 in enumerate(files5):
    if file5 != "temp.docx":
        source_file5 = os.path.join(source_directory5, file5)
        target_file5 = os.path.join(target_directory5, file5)
        shutil.move(source_file5, target_file5)

# 移动 ./template/minor/catalog15 目录下的文件
files6 = os.listdir(source_directory6)
for file6 in files6:
    if file6 != "temp.docx":
        source_file6 = os.path.join(source_directory6, file6)
        target_file6 = os.path.join(target_directory6, file6)
        shutil.move(source_file6, target_file6)

# 删除 ../output/doc/移交签字 目录下的除第一个文件外的全部文件
files_to_keep = files5[:1]
files_in_target_directory5 = os.listdir(target_directory5)
for file_in_target_directory5 in files_in_target_directory5:
    if file_in_target_directory5 not in files_to_keep:
        file_path = os.path.join(target_directory5, file_in_target_directory5)
        os.remove(file_path)

# 删除目录及其子目录下的所有 export.xlsx 文件
def delete_export_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file == "export.xlsx":
                file_path = os.path.join(root, file)
                os.remove(file_path)

delete_export_files(directory_path)