@echo off
title CETEM Flow — RDA Professional
color 1F
echo.
echo  ==========================================
echo   CETEM Flow — iniciando servidor local...
echo  ==========================================
echo.

cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo  [AVISO] Ambiente virtual nao encontrado. Criando...
    python -m venv .venv
    echo  [OK] Ambiente virtual criado.
    echo.
    echo  Instalando dependencias...
    call .venv\Scripts\activate.bat
    pip install -r requirements.txt
) else (
    call .venv\Scripts\activate.bat
)

echo  Ambiente virtual ativo.
echo.
echo  Acesse: http://localhost:5000
echo  Para encerrar: pressione Ctrl+C
echo.

python app.py

pause
