@echo off
rem Doble clic para preparar esta computadora para los plugins de Claude del Estudio BVA.
rem Se puede correr las veces que haga falta.
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar.ps1" %*
echo.
pause
