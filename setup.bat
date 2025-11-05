@echo off
chcp 65001 >nul
echo ========================================
echo ITC AutoReminder 项目初始化脚本
echo ========================================
echo.

:: 检查 Python 是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8 或更高版本
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/6] 检测到 Python 版本:
python --version
echo.

:: 创建虚拟环境
echo [2/6] 创建 Python 虚拟环境...
if not exist ".venv" (
    python -m venv .venv
    echo ✓ 虚拟环境创建成功
) else (
    echo ✓ 虚拟环境已存在
)
echo.

:: 激活虚拟环境并安装依赖
echo [3/6] 安装 Python 依赖包...
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [错误] 依赖包安装失败，请检查网络连接或 requirements.txt 文件
    pause
    exit /b 1
)
echo ✓ 依赖包安装成功
echo.

:: 创建必要的文件夹
echo [4/6] 创建项目文件夹结构...
if not exist "ITC report" mkdir "ITC report"
if not exist "ITC report\RawData" mkdir "ITC report\RawData"
if not exist "ITC report\Reminder" mkdir "ITC report\Reminder"
echo ✓ 文件夹创建完成
echo.

:: 创建配置文件
echo [5/6] 生成配置文件模板...

:: 检查并创建 email_config.json
if not exist "email_config.json" (
    if exist "email_config.json.example" (
        copy "email_config.json.example" "email_config.json" >nul
        echo ✓ 已创建 email_config.json，请修改其中的邮箱地址
    ) else (
        echo ! 未找到 email_config.json.example，跳过
    )
) else (
    echo ✓ email_config.json 已存在
)

:: 检查并创建 teams_config.json
if not exist "teams_config.json" (
    if exist "teams_config.json.example" (
        copy "teams_config.json.example" "teams_config.json" >nul
        echo ✓ 已创建 teams_config.json，请修改其中的 Webhook URL
    ) else (
        echo ! 未找到 teams_config.json.example，跳过
    )
) else (
    echo ✓ teams_config.json 已存在
)
echo.

:: 完成
echo [6/6] 初始化完成！
echo.
echo ========================================
echo 下一步操作：
echo ========================================
echo 1. 编辑配置文件：
echo    - email_config.json （配置邮箱地址）
echo    - teams_config.json  （配置 Teams Webhook）
echo.
echo 2. 运行程序：
echo    - 方式一：python BatRun_ITCreport_downloader_rev1.py
echo    - 方式二：双击 Run_OneDriveITC.bat
echo.
echo 3. 首次运行需要手动登录 Chrome 中的 ITC 系统
echo.
echo 详细说明请查看 SETUP.md 文档
echo ========================================
echo.
pause
