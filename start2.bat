@echo off

rem 运行06_changesheet1.py
python 06_changesheet1.py

rem 检查06_changesheet1.py的返回代码
if %errorlevel% neq 0 (
    echo "06_changesheet1.py 运行出错"
    exit /b
)

rem 运行07_changesheet2.py
python 07_changesheet2.py

rem 检查07_changesheet2.py的返回代码
if %errorlevel% neq 0 (
    echo "07_changesheet2.py 运行出错"
    exit /b
)

rem 运行08_changelittle.py
python 08_changelittle.py

rem 检查08_changelittle.py的返回代码
if %errorlevel% neq 0 (
    echo "08_changelittle.py 运行出错"
    exit /b
)

rem 运行09_mergerxls1.py
python 09_mergerxls1.py

rem 检查09_mergerxls1.py的返回代码
if %errorlevel% neq 0 (
    echo "09_mergerxls1.py 运行出错"
    exit /b
)

rem 运行10_mergerxls2.py
python 10_mergerxls2.py

rem 检查10_mergerxls2.py的返回代码
if %errorlevel% neq 0 (
    echo "10_mergerxls2.py 运行出错"
    exit /b
)

rem 运行11_copydel.py
python 11_copydel.py

rem 检查11_copydel.py的返回代码
if %errorlevel% neq 0 (
    echo "11_copydel.py 运行出错"
    exit /b
)

rem 结束运行
exit /b
