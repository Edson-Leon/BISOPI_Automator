@echo off
echo ============================================
echo   BISOPI Automator — Instalacion
echo ============================================
echo.
echo Verificando Python...
python --version
if errorlevel 1 (
    echo ERROR: Python no encontrado.
    echo Instala Python 3.10 o superior desde https://www.python.org
    pause
    exit /b 1
)
echo.
echo Instalando dependencias...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Fallo la instalacion de dependencias.
    pause
    exit /b 1
)
echo.
echo ============================================
echo   Instalacion completa.
echo ============================================
echo.
echo Proximos pasos:
echo   1. Copia .env.example a .env
echo   2. Edita .env con tus credenciales
echo   3. Ejecuta run.bat para arrancar la app
echo.
pause
