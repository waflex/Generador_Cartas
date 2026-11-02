@echo off
setlocal
cd /d "%~dp0"
where pythonw >nul 2>nul
if errorlevel 1 (
    python "%~dp0interfaz_cartas.py"
) else (
    start "" pythonw "%~dp0interfaz_cartas.py"
)
endlocal
