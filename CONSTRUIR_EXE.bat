@echo off
setlocal
title JAC Nuevo Achi - CONSTRUCTOR EXE
cd /d "%~dp0"
where py >nul 2>&1
if %errorlevel%==0 (set "PY=py") else (
  where python >nul 2>&1
  if %errorlevel%==0 (set "PY=python") else (
    echo Python no esta instalado.
    echo Instala Python 3.11+ y vuelve a ejecutar este archivo.
    pause
    exit /b 1
  )
)
echo.
echo Instalando herramientas...
%PY% -m pip install Flask==3.1.2 openpyxl==3.1.5 pyinstaller==6.16.0
if errorlevel 1 goto error
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist JAC_Nuevo_Achi.spec del /q JAC_Nuevo_Achi.spec
echo.
echo Construyendo JAC_Nuevo_Achi.exe...
%PY% -m PyInstaller --noconfirm --clean --onefile --windowed --name JAC_Nuevo_Achi --add-data "templates;templates" --add-data "static;static" --add-data "data;data" launcher.py
if errorlevel 1 goto error
echo.
echo ==========================================================
echo EXE CREADO CORRECTAMENTE:
echo %CD%\dist\JAC_Nuevo_Achi.exe
echo ==========================================================
echo.
echo Copia ese EXE al PC donde quieras usarlo.
pause
exit /b 0
:error
echo.
echo ERROR: No se pudo construir el EXE.
echo Revisa la conexion a Internet y que Python sea 3.11+.
pause
exit /b 1
