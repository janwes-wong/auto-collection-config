@echo off

echo =====***** 欢迎使用交换机配置自动采集工具 *****=====
:: 设置项目根路径
set "PROJECT_ROOT=%~dp0"
echo 项目路径： %PROJECT_ROOT%

:: 设置虚拟环境路径
set "VENV_DIR=%PROJECT_ROOT%.venv\Scripts\activate.bat"
echo 虚拟环境路径： %VENV_DIR%

:: 检查虚拟环境是否存在
if not exist "%VENV_DIR%" (
    echo 错误：未找到虚拟环境
    echo 路径: %VENV_DIR%
    pause
    exit /b 1
)

echo =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*

:: 激活环境并运行
call "%VENV_DIR%"
cd /d "%PROJECT_ROOT%"
python main.py

:: 检查脚本是否成功运行
if %errorlevel% equ 0 (
    echo 本次任务执行完成！
) else (
    echo 脚本执行失败，错误码：%errorlevel%
)

pause