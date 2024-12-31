@echo off
setlocal enabledelayedexpansion

:: ��ȡ��ǰ·��
set CURRENT_PATH=%cd%

:: ��ȡ�û��������� Path
for /f "tokens=2,*" %%A in ('reg query HKCU\Environment /v Path 2^>nul') do set USER_PATH=%%B

:: �жϵ�ǰ·���Ƿ��� Path ��
echo !USER_PATH! | findstr /i /c:"%CURRENT_PATH%" >nul
if not errorlevel 1 (
    echo ��ǰλ�����ڻ��������У������ٴΰ�װ��
    pause
    exit /b
)

:: ��ʾ�û�ȷ�ϰ�װ
set /p CONFIRM=[��װ NAVI] ���������ȷ�ϰ�װ NAVI ...

:: �ж��Ƿ��ԷֺŽ�β
set LAST_CHAR=!USER_PATH:~-1!
if "!LAST_CHAR!"==";" (
    set NEW_PATH=!USER_PATH!!CURRENT_PATH!
) else (
    set NEW_PATH=!USER_PATH!;%CURRENT_PATH%
)

:: �޸�ע��������û��������� Path
reg add HKCU\Environment /v Path /t REG_EXPAND_SZ /d "!NEW_PATH!" /f
setx Path "!NEW_PATH!"

:: ��ʾ��װ���
echo ���ڣ��������������� "NAVI" ��ʱ���� NAVI �ˡ�
echo.
pause