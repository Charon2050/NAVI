@echo off
setlocal enabledelayedexpansion

:: 获取当前路径
set "CURR_PATH=%cd%"

:: 检查当前路径是否在用户环境变量Path中
for /f "tokens=*" %%i in ('reg query HKCU\Environment /v Path ^| findstr Path') do set "USER_PATH=%%i"

:: 提取Path内容
set "USER_PATH=!USER_PATH:*REG_SZ    =!"

:: 判断路径是否已存在
echo !USER_PATH! | findstr /i "%CURR_PATH%" >nul
if not errorlevel 1 (
    echo 当前位置已在环境变量中，无需再次安装。
    pause
    exit /b
)

:: 询问用户是否添加
echo [安装 NAVI] 按下任意键确认安装 NAVI ...
pause >nul

echo.
:: 添加路径到用户环境变量
set "NEW_PATH=!USER_PATH!;%CURR_PATH%"

:: 修改注册表以更新Path
reg add HKCU\Environment /v Path /t REG_SZ /d "%NEW_PATH%" /f

:: 提示安装成功
echo 现在，您可以用命令行“NAVI”随时启动 NAVI 了。
pause