@echo off
:: Establece la ruta de trabajo a la ubicación del archivo .bat
cd /d %~dp0

:: Comprueba si se está ejecutando como administrador
net session >nul 2>&1
if %errorLevel% NEQ 0 (
    echo Requiere permisos de administrador. Ejecutando de nuevo con privilegios elevados...
    powershell start-process "%~0" -verb runas
    exit
)

:: Ejecuta el comando regsvr32 para registrar la DLL
regsvr32 DX8VB.DLL
