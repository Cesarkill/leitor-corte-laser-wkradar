@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist ".venv\Scripts\activate.bat" (
    call ".venv\Scripts\activate.bat"
) else (
    echo [AVISO] Ambiente virtual .venv nao encontrado. Usando o Python do sistema.
)

taskkill /F /IM LeitorCorteLaser.exe >nul 2>&1
if not errorlevel 1 timeout /t 1 /nobreak >nul

pyinstaller --clean --noconfirm --onefile --windowed --name "LeitorCorteLaser" main.py
if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao gerar o executavel! Feche o LeitorCorteLaser.exe e tente novamente.
    pause
    exit /b 1
)

if exist ".env" copy /Y ".env" "dist\.env" >nul

echo.
echo Executavel gerado em: dist\LeitorCorteLaser.exe
echo Lembre-se: o arquivo .env com a GEMINI_API_KEY precisa ficar na mesma pasta do .exe.
pause
