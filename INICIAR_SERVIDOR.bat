@echo off
title ERP CONSITEC / Poderosa - Servidor Local
cd /d "%~dp0"
echo Iniciando servidor del ERP CONSITEC...
echo.
powershell -NoExit -ExecutionPolicy Bypass -File "%~dp0sharepoint\iniciar-servidor.ps1"
