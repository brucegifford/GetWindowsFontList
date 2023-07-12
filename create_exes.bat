@echo off
cd /D %~dp0

REM THIS SCRIPT HAS THE FOLLOWING DEPENDENCY
REM pip install pyinstaller

IF EXIST build (
	rmdir build /s/q
)
IF EXIST dist (
	rmdir dist /s/q
)

call venv\Scripts\activate.bat
pyinstaller.exe --onefile GetWindowsFontList.py
