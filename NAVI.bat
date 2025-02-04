@echo off
cd /d "%~dp0"
set path=%cd%\python;%cd%\tools;%path%;
python "NAVI.py" %*
echo --------------------
echo 很抱歉，程序遇到问题，需要退出。请查看以上报错信息，或发给开发者以便修复。
pause