@echo off
setlocal enabledelayedexpansion

:: ��ȡ��ǰ·��
set "CURR_PATH=%cd%"

:: ��鵱ǰ·���Ƿ����û���������Path��
for /f "tokens=*" %%i in ('reg query HKCU\Environment /v Path ^| findstr Path') do set "USER_PATH=%%i"

:: ��ȡPath����
set "USER_PATH=!USER_PATH:*REG_SZ    =!"

:: �ж�·���Ƿ��Ѵ���
echo !USER_PATH! | findstr /i "%CURR_PATH%" >nul
if not errorlevel 1 (
    echo ��ǰλ�����ڻ��������У������ٴΰ�װ��
    pause
    exit /b
)

:: ѯ���û��Ƿ����
echo [��װ NAVI] ���������ȷ�ϰ�װ NAVI ...
pause >nul

echo.
:: ���·�����û���������
set "NEW_PATH=!USER_PATH!;%CURR_PATH%"

:: �޸�ע����Ը���Path
reg add HKCU\Environment /v Path /t REG_SZ /d "%NEW_PATH%" /f

:: ��ʾ��װ�ɹ�
echo ���ڣ��������������С�NAVI����ʱ���� NAVI �ˡ�
pause