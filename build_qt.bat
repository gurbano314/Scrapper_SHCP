@echo off
REM ============================================================
REM  build_qt.bat  –  Compilador para Conciliador v6 (PyQt6)
REM  Crea un .exe nativo sin Streamlit ni servidor web.
REM ============================================================
setlocal EnableDelayedExpansion

set VENV=build_venv_qt
set PYTHON=python
set PYVER_REQ=3.12

echo.
echo ============================================================
echo   Conciliador v6 ^| Build Qt  ^|  %DATE% %TIME%
echo ============================================================

REM -- Verificar Python ----------------------------------------
%PYTHON% --version >nul 2>&1
if errorlevel 1 (
    echo [Error] Python no encontrado en PATH.
    echo         Instala Python %PYVER_REQ% desde python.org
    exit /b 1
)

REM -- Crear / activar virtualenv ------------------------------
if not exist %VENV% (
    echo [1/5] Creando entorno virtual %VENV%...
    %PYTHON% -m venv %VENV%
    if errorlevel 1 ( echo [Error] Fallo al crear venv & exit /b 1 )
) else (
    echo [1/5] Usando entorno existente %VENV%
)

call %VENV%\Scripts\activate.bat

REM -- Actualizar pip ------------------------------------------
echo [2/5] Actualizando pip...
python -m pip install --upgrade pip

REM -- Instalar dependencias -----------------------------------
echo [3/5] Instalando dependencias...
pip install -r requirements_qt.txt
if errorlevel 1 ( echo [Error] Fallo al instalar dependencias & exit /b 1 )

REM -- Limpiar artefactos anteriores ---------------------------
echo [4/5] Limpiando builds anteriores...
if exist dist\Conciliador rmdir /s /q dist\Conciliador
if exist build\Conciliador rmdir /s /q build\Conciliador

REM -- Compilar ------------------------------------------------
echo [5/5] Compilando con PyInstaller...
pyinstaller conciliador_qt.spec --noconfirm --clean
if errorlevel 1 ( echo [Error] Fallo en PyInstaller & exit /b 1 )

echo.
echo ============================================================
echo   Build exitoso: dist\Conciliador\Conciliador.exe
echo ============================================================
echo.

REM -- Abrir carpeta de salida ---------------------------------
if exist dist\Conciliador (
    explorer dist\Conciliador
)


