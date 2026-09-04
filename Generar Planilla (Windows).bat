@echo off
REM ============================================================
REM   GENERADOR DE PLANILLA BISEMANAL - Ingenieria Estrella S.A.
REM   Doble clic para abrir. Se abre solo en el navegador.
REM ============================================================
setlocal enabledelayedexpansion
cd /d "%~dp0"
title Generador de Planilla Bisemanal
mode con: cols=88 lines=26

REM ---------- falta la libreria incluida ----------
if not exist "programa\lib\openpyxl" (
    call :AVISO "Falta la carpeta programa\lib.`n`nEsa carpeta trae la libreria que lee los archivos de Excel.`n`nCopie de nuevo la carpeta completa del programa."
    echo.
    echo   FALTA LA CARPETA  programa\lib
    echo.
    pause
    exit /b 1
)

REM ---------- buscar Python ----------
set PY=
for %%C in (py python python3) do (
    if not defined PY (
        %%C -c "" >nul 2>&1 && set PY=%%C
    )
)

if not defined PY (
    call :PREGUNTA
    echo.
    echo   Falta instalar Python. Cuando termine, vuelva a dar doble clic aqui.
    echo.
    pause
    exit /b 1
)

REM ---------- todo listo ----------
echo.
echo   Guia de uso:  programa\LEEME.txt
echo.
%PY% "programa\servidor.py"
pause
exit /b 0

REM ============================================================
:AVISO
powershell -NoProfile -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show(%1, 'Generador de Planilla Bisemanal', 'OK', 'Error')" >nul 2>&1
exit /b

:PREGUNTA
set MSG="Esta computadora todavia no tiene Python, el motor que necesita el programa.`n`nSe instala UNA sola vez, es gratis y viene del sitio oficial python.org.`n`nMUY IMPORTANTE: durante la instalacion marque la casilla 'Add python.exe to PATH' antes de presionar Install.`n`nQuiere abrir la pagina de descarga ahora?"
for /f %%R in ('powershell -NoProfile -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show(%MSG%, 'Generador de Planilla Bisemanal', 'YesNo', 'Warning')" 2^>nul') do set RESP=%%R
if /i "%RESP%"=="Yes" (
    start "" "https://www.python.org/downloads/"
    powershell -NoProfile -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Se abrio la pagina de descarga en el navegador.`n`nRecuerde marcar la casilla ''Add python.exe to PATH'' durante la instalacion.`n`nAl terminar, vuelva a dar doble clic en este archivo.', 'Generador de Planilla Bisemanal', 'OK', 'Information')" >nul 2>&1
) else (
    powershell -NoProfile -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Cuando quiera hacerlo, entre a https://www.python.org/downloads/ y descargue Python.`n`nMarque ''Add python.exe to PATH'' durante la instalacion.`n`nAl terminar, vuelva a dar doble clic en este archivo.', 'Generador de Planilla Bisemanal', 'OK', 'Information')" >nul 2>&1
)
exit /b
