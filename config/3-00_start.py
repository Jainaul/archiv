import os
import subprocess

# 运行每个 .py 文件的列表
script_files = [
    "3-01_generatedoc1.py",
    "3-02_generatedoc2.py",
    "3-03_generatedoc3.py",
    "3-04_generatedoc4.py",
    "3-05_generatedocx1.py",
    "3-06_generatedocx2.py",
    "3-07_moverm.py"
]

# 创建窗口
print("===================================")
print("程序开始运行")
print("===================================")

error_occurred = False

# 依次运行每个 .py 文件
for script in script_files:
    # 显示正在运行的文件名
    print(f"\n{script} 正在运行")

    # 运行 .py 文件
    try:
        subprocess.check_call(["python", script])
        print(f"{script} 运行完成")
    except subprocess.CalledProcessError:
        print("出错了，请检查PDF源文件")
        error_occurred = True
        break  # 出错则停止继续运行其他文件

# 完成（仅当没有出错时才打印）
if not error_occurred:
    print("\n===================================")
    print("运行完成，请核对生成的文件")
    print("===================================")