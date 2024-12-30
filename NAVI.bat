@echo off
cd /d "%~dp0"
set path=%cd%\python;%cd%\tools;%path%;
python "NAVI.py" %*