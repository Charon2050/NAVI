@echo off
cd /d "%~dp0"
set path=%cd%\python;%cd%\tools;%path%;
python "NAVI.py" %*
echo --------------------
echo �ܱ�Ǹ�������������⣬��Ҫ�˳�����鿴���ϱ�����Ϣ���򷢸��������Ա��޸���
pause