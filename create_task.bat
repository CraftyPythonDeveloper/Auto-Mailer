@echo off

set "PYTHON_PATH=C:\Users\%USERNAME%\anaconda3\python.exe"
set "SCRIPT_PATH=%~dp0\src\task_scheduler.py"
call %PYTHON_PATH% %SCRIPT_PATH%
