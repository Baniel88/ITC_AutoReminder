@echo off
cd /d %~dp0
:: 使用本地安装的Python313路径（更稳定）
set python_path="C:\Users\liang.wq.1\AppData\Local\Programs\Python\Python313\python.exe"
:: 执行下载脚本
echo 正在执行下载脚本...
python BatRun_ITCreport_downloader.py
IF ERRORLEVEL 1 (
    echo 下载脚本执行失败，停止执行。
    pause
    exit /b 1
)
echo 下载脚本执行完毕。

:: 执行 pending_review_report.py
echo 正在执行 pending_review_report.py...
python pending_review_report.py
IF ERRORLEVEL 1 (
    echo pending_review_report.py 执行失败，停止执行。
    pause
    exit /b 1
)
echo pending_review_report.py 执行完毕。

:: 执行 email_sender.py
echo 正在执行 email_sender.py...
python email_sender.py
IF ERRORLEVEL 1 (
    echo email_sender.py 执行失败。
    pause
    exit /b 1
)
echo email_sender.py 执行完毕。

pause