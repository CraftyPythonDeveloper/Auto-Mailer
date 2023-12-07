@echo off
set /p TASKNAME="Enter the task name: "
schtasks /delete /tn "%TASKNAME%" /f
pause
