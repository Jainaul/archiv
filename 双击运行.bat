@echo off
call "%~dp0/wpy64/scripts/env_for_icons.bat"  %*
if not "%WINPYWORKDIR%"=="%WINPYWORKDIR1%" cd %WINPYWORKDIR1%
cd /d "%~dp0\config"
cmd.exe /k python main.py
