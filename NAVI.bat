@echo off
cd /d "%~dp0"
set path=%cd%\python;%path%;
python "NAVI.py" %*