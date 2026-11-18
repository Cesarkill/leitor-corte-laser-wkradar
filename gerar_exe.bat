@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist ".venv\Scripts\activate.bat" (
    call ".venv\Scripts\activate.bat"
) else (
    echo [AVISO] Ambiente virtual .venv nao encontrado. Usando o Python do sistema.
)

pyinstaller --onefile --windowed --name "LeitorCorteLaser" main.py

if exist ".env" copy /Y ".env" "dist\.env" >nul

echo.
echo Executavel gerado em: dist\LeitorCorteLaser.exe
echo Lembre-se: o arquivo .env com a GEMINI_API_KEY precisa ficar na mesma pasta do .exe.
pause
