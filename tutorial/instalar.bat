@echo off
setlocal
chcp 65001 >nul
title Instalador - GLP (Quintana Energy)

set "ADDIN_DIR=%USERPROFILE%\AppData\Roaming\Microsoft\AddIns"
set "MANIFEST=%ADDIN_DIR%\glp-manifest.prod.xml"
set "MANIFEST_URL=https://github.com/jleal-quintana/GLP/raw/main/manifest.prod.xml"

echo.
echo ============================================================
echo   GLP - Instalador para Excel
echo   Quintana Energy
echo ============================================================
echo.

echo [1/3] Preparando carpeta de complementos...
if not exist "%ADDIN_DIR%" mkdir "%ADDIN_DIR%"

echo [2/3] Descargando configuracion (manifest)...
powershell -NoProfile -Command "try { Invoke-WebRequest -UseBasicParsing -Uri '%MANIFEST_URL%' -OutFile '%MANIFEST%' -ErrorAction Stop; exit 0 } catch { exit 1 }"

if not exist "%MANIFEST%" (
    echo.
    echo ERROR: No se pudo descargar el manifest.
    echo Verifica que tengas conexion a internet y volve a intentar.
    echo.
    pause
    exit /b 1
)

echo [3/3] Registrando complemento en Excel...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "GLP" /t REG_SZ /d "%MANIFEST%" /f >nul

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: No se pudo registrar el complemento.
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   LISTO - Instalacion completada
echo ============================================================
echo.
echo Proximos pasos:
echo.
echo  1. Cerrar Excel completamente si lo tenes abierto.
echo  2. Volver a abrir Excel con un libro nuevo.
echo  3. Pestana HOME/Inicio - boton ADD-INS al final de la cinta.
echo  4. En el menu, abajo, click en MORE ADD-INS.
echo  5. Tab DEVELOPER ADD-INS - click en GLP.
echo  6. El panel se abre a la derecha.
echo.
pause
