@echo off
setlocal enabledelayedexpansion

:: 获取当前路径
set CURRENT_PATH=%cd%

:: 获取用户环境变量 Path
for /f "tokens=2,*" %%A in ('reg query HKCU\Environment /v Path 2^>nul') do set USER_PATH=%%B

:: 判断当前路径是否在 Path 中
echo !USER_PATH! | findstr /i /c:"%CURRENT_PATH%" >nul
if not errorlevel 1 (
    echo 当前位置已在环境变量中，无需再次安装。
    pause
    exit /b
)

:: 提示用户确认安装
set /p CONFIRM=[安装 NAVI] 按下任意键确认安装 NAVI ...

:: 判断是否以分号结尾
set LAST_CHAR=!USER_PATH:~-1!
if "!LAST_CHAR!"==";" (
    set NEW_PATH=!USER_PATH!!CURRENT_PATH!
) else (
    set NEW_PATH=!USER_PATH!;%CURRENT_PATH%
)

:: 修改注册表，更新用户环境变量 Path
reg add HKCU\Environment /v Path /t REG_EXPAND_SZ /d "!NEW_PATH!" /f
setx Path "!NEW_PATH!"

:: 提示安装完成
echo 现在，您可以用命令行 "NAVI" 随时启动 NAVI 了。
echo.
pause