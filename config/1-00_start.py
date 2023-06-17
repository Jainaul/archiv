import os
import subprocess

# 运行每个 .py 文件的列表
script_files = [
    "1-01_creatsheet.py",
    "1-02_addbadge.py",
    "1-03_addnum.py",
    "1-04_addno.py",
    "1-05_generatepage.py"
]

# 创建窗口
print("===================================")
print("程序开始运行")
print("===================================")

# 询问用户是否显示详细信息
show_details = input("是否显示详细信息？ (y/n): ").lower()

error_occurred = False

# 依次运行每个 .py 文件
for script in script_files:
    # 显示正在运行的文件名
    print(f"\n{script} 正在运行")

    # 运行 .py 文件
    try:
        if show_details == "n":
            subprocess.check_output(["python", script], stderr=subprocess.STDOUT)
        else:
            subprocess.check_call(["python", script])
        print(f"{script} 运行完成")
    except subprocess.CalledProcessError:
        print("出错了，请检查PDF源文件")
        error_occurred = True
        break  # 出错则停止继续运行其他文件

# 删除 output.xlsx 文件
print("\n正在清除临时文件")
file_path = "./output.xlsx"
if os.path.exists(file_path):
    os.remove(file_path)
    print("output.xlsx 已成功删除")

# 完成（仅当没有出错时才打印）
if not error_occurred:
    print("\n===================================")
    print("运行完成，请核对生成的文件")
    print("===================================")
