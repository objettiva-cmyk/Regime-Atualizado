@echo off
REM ============================================================
REM Motor Fiscal RIR70 by MC4
REM Proprietario: MC4 CONTABILIDADE E GESTAO DE NEGOCIOS
REM CNPJ 09.944.432/0001-25
REM Ferramenta proprietaria licenciada para uso interno
REM ============================================================
title MC4 - Motor Fiscal RIR70 by MC4 - Sistema
cd /d "%~dp0"

echo ============================================================
echo MC4 - MOTOR FISCAL RIR70 BY MC4 - SISTEMA
echo ============================================================
echo Pasta do sistema: %CD%
echo.

if not exist "sistema_rir70_web.py" (
    echo [ERRO] sistema_rir70_web.py nao localizado.
    pause
    exit /b 1
)

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 sistema_rir70_web.py
) else (
    python sistema_rir70_web.py
)

pause
