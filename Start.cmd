@echo off
setlocal
set "TEX2WORD_PYTHON=%~dp0.venv\Scripts\python.exe"
set "PYTHONPATH=%~dp0;%PYTHONPATH%"
if not exist "%TEX2WORD_PYTHON%" goto setup
"%TEX2WORD_PYTHON%" -c "import importlib.util as u, sys; sys.exit(not all(u.find_spec(n) for n in ('rapid_doc','win32com','docx','pypandoc','olefile')))"
if not errorlevel 1 goto run
:setup
call "%~dp0Setup.cmd"
if errorlevel 1 exit /b 1
:run
if "%~1"=="" (
    start "" "%~dp0.venv\Scripts\pythonw.exe" -m tex2word
) else (
    "%TEX2WORD_PYTHON%" -m tex2word %*
    if errorlevel 1 exit /b 1
)
exit /b 0
