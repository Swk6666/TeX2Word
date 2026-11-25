@echo off
setlocal
pushd "%~dp0"
where uv >nul 2>nul
if errorlevel 1 goto python
if not exist ".venv\Scripts\python.exe" (
    uv venv --python 3.12 .venv
    if errorlevel 1 goto failed
)
uv pip install --python ".venv\Scripts\python.exe" -r requirements.txt
if errorlevel 1 goto failed
goto ready
:python
if exist ".venv\Scripts\python.exe" goto install
where py >nul 2>nul
if errorlevel 1 (
    python -m venv .venv
) else (
    py -3 -m venv .venv
)
if errorlevel 1 goto failed
:install
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
:ready
echo TeX2Word is ready. Open Start.cmd.
popd
exit /b 0
:failed
echo Setup failed. Install Python 3.12 x64 or uv, then run Setup.cmd again.
pause
exit /b 1
