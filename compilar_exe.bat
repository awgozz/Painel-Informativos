@echo off
cd /d "%~dp0"

python -m PyInstaller --clean --noconfirm PainelInformativosGarbuio.spec

if errorlevel 1 (
    echo.
    echo Nao foi possivel gerar o executavel.
    pause
    exit /b 1
)

echo.
echo Executavel gerado em:
echo %~dp0dist\PainelInformativosGarbuio.exe
pause
