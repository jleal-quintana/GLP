@echo off
setlocal
chcp 65001 >nul
title Desinstalador - GLP (Quintana Energy)

set "ADDIN_DIR=%USERPROFILE%\AppData\Roaming\Microsoft\AddIns"
set "MANIFEST=%ADDIN_DIR%\glp-manifest.prod.xml"

echo.
echo ============================================================
echo   GLP - Desinstalador para Excel
echo ============================================================
echo.

echo Quitando registro del complemento...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "GLP" /f >nul 2>nul

echo Quitando manifest local...
if exist "%MANIFEST%" del "%MANIFEST%"

echo.
echo LISTO. Cerrar Excel completamente y volver a abrirlo.
echo.
pause
