@echo off

set findpip=0

@for /f "tokens=1" %%i in ('pip --version ^| findstr /C:"pip"') do ^
if %%i == pip set findpip=1

if  %findpip% == 1 (
	python -m pip --version
) else (
	echo "need download pip"
	rem python downpip.py
	python -m pip install -U pip
)
  
pip install xlrd
python exls2file.py

pause