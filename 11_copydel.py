import os
import shutil

# 复制文件
source_file = "./outputchange.xlsx"
destination_folder = "./main"
destination_file = "./main/export.xlsx"

shutil.copyfile(source_file, destination_file)
print(f"文件已成功复制到 {destination_file}")

# 删除源文件和目标文件
os.remove(source_file)
os.remove("./output.xlsx")
print(f"源文件 {source_file} 和目标文件 ./output.xlsx 已删除")

# 删除archiv文件夹中的文件
archive_directory = "./archiv"
for root, dirs, files in os.walk(archive_directory):
    for file in files:
        if file == "outputchange.xlsx" or file == "output.xlsx":
            file_path = os.path.join(root, file)
            os.remove(file_path)
            print(f"已删除文件: {file_path}")
