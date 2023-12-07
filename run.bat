@echo off

set "PYTHON_PATH=C:\Users\%USERNAME%\anaconda3\python.exe"
set "SCRIPT_PATH=%~dp0\src\auto_mailer.py"
set /p GROUP_ID="Enter the group_id: "
call %PYTHON_PATH% %SCRIPT_PATH% %GROUP_ID%
