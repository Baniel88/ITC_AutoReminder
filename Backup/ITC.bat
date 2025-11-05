
@echo off
:: 切换到脚本所在目录（关键！）
cd /d %~dp0
:: 使用本地安装的Python313路径（更稳定）
set python_path="C:\Users\liang.wq.1\AppData\Local\Programs\Python\Python313\python.exe"
set python_file_path="C:\Users\liang.wq.1\Downloads\ITC_Scorecard\BatRun_ITCreport_downloader.py"

:: 执行脚本并显示可能的错误信息
%python_path% %python_file_path%

:: 设置3小时后自动退出
timeout /t 10800 /nobreak >nul

:: 如果执行失败，会停留在此界面显示错误
pause