@echo off

REM 激活 Anaconda 环境（如果需要）
call activate base

REM 安装模块
conda install --channel https://mirrors.aliyun.com/anaconda/pkgs/main/win-64/ openpyxl PyPDF2 reportlab pywin32 && conda install -c conda-forge pyautogui

REM 检查安装是否成功
conda list | findstr openpyxl
conda list | findstr PyPDF2
conda list | findstr reportlab
conda list | findstr pyautogui
conda list | findstr pywin32

REM 暂停脚本执行，以便查看输出
pause
