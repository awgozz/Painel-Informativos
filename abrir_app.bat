@echo off
cd /d "%~dp0"
python app.py

if errorlevel 1 (
    echo.
    echo Nao foi possivel abrir o app.
    echo Confira se o Python esta instalado e tente novamente.
    pause
)
