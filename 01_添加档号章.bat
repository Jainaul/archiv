@echo off
chcp 65001 > nul

rem 切换到"./config"文件夹
cd /d "%~dp0config"

rem 运行1-00_start.py
python 1-00_start.py

rem 检查1-00_start.py的返回代码
if %errorlevel% neq 0 (
    echo "1-00_start.py 运行出错"
    pause
    exit /b
)

rem 结束运行
echo 输入空格键关闭本窗口
echo ===================================
pause > nul
exit /b
