@echo off

rem 运行01_creatsheet.py
python 01_creatsheet.py

rem 检查01_creatsheet.py的返回代码
if %errorlevel% neq 0 (
    echo "01_creatsheet.py 运行出错"
    exit /b
)

rem 运行02_addbadge.py
python 02_addbadge.py

rem 检查02_addbadge.py的返回代码
if %errorlevel% neq 0 (
    echo "02_addbadge.py 运行出错"
    exit /b
)

rem 运行03_addnum.py
python 03_addnum.py

rem 检查03_addnum.py的返回代码
if %errorlevel% neq 0 (
    echo "03_addnum.py 运行出错"
    exit /b
)

rem 运行04_addno.py
python 04_addno.py

rem 检查04_addno.py的返回代码
if %errorlevel% neq 0 (
    echo "04_addno.py 运行出错"
    exit /b
)

rem 运行05_generatepage.py
python 05_generatepage.py

rem 检查05_generatepage.py的返回代码
if %errorlevel% neq 0 (
    echo "05_generatepage.py 运行出错"
    exit /b
)

rem 结束运行
exit /b
