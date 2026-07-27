@echo off
REM Launch the Royal Chain HR Dashboard (Streamlit)
setlocal

set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not exist "%PYEXE%" set "PYEXE=python"

cd /d "%~dp0"
"%PYEXE%" -m streamlit run app.py
endlocal
