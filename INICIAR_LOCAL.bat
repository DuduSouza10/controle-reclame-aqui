@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo Dashboard Reclame Aqui - J^&T Express
echo ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
  echo Python nao foi encontrado no PATH.
  pause
  exit /b 1
)

echo Instalando/verificando dependencias...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Falha ao instalar dependencias.
  pause
  exit /b 1
)

echo.
echo Abrindo dashboard em http://127.0.0.1:5000
python app.py
pause
