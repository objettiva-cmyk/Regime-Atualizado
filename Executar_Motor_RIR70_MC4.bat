@echo off
setlocal EnableExtensions EnableDelayedExpansion
REM ============================================================
REM MC4 - MOTOR FISCAL RIR70 BY MC4 - PRODUCAO
REM Execucao unica com pre-validacao fail-fast no motor.
REM ============================================================

title MC4 - Motor Fiscal RIR70 by MC4 - Producao
set "ROOT=%~dp0"
cd /d "%ROOT%"

if not exist "%ROOT%logs_integridade\producao" mkdir "%ROOT%logs_integridade\producao" >nul 2>nul
set "LOG_DEP=%ROOT%logs_integridade\producao\dependencias_mc4.log"

if "%MC4_LAYOUT_NIVEL%"=="" set "MC4_LAYOUT_NIVEL=auditoria"

echo ============================================================
echo MC4 - MOTOR FISCAL RIR70 BY MC4 - PRODUCAO
echo ============================================================
echo Pasta do sistema: %ROOT%
echo Layout: %MC4_LAYOUT_NIVEL%
echo Regra: qualquer falha de integridade ou input bloqueia o calculo.
echo.

set "PYTHON_EXE="
where py >nul 2>nul
if not errorlevel 1 set "PYTHON_EXE=py -3"
if "%PYTHON_EXE%"=="" (
    where python >nul 2>nul
    if not errorlevel 1 set "PYTHON_EXE=python"
)
if "%PYTHON_EXE%"=="" (
    echo [ERRO] Python nao localizado.
    goto FIM_ERRO
)

for %%F in (motor_arbitramento.py layout_premium_mc4.py config_arbitramento_rir70.json VERSAO_MC4.json MANIFESTO_PACOTE_MC4.json MANIFESTO_PACOTE_MC4.sig MANIFESTO_PACOTE_MC4.sha256 MC4_PUBLIC_KEY.pem) do (
    if not exist "%ROOT%%%F" (
        echo [ERRO] Arquivo critico ausente: %%F
        goto FIM_ERRO
    )
)

echo [OK] Python localizado.
%PYTHON_EXE% --version

echo Verificando dependencias...
%PYTHON_EXE% -c "import cryptography, flask, openpyxl, pandas, xlsxwriter; import python_calamine" >nul 2>>"%LOG_DEP%"
if errorlevel 1 (
    if not exist "%ROOT%requirements.txt" (
        echo [ERRO] requirements.txt ausente e dependencias nao instaladas.
        goto FIM_ERRO
    )
    echo [INFO] Instalando dependencias. Aguarde...
    %PYTHON_EXE% -m pip install -q -r "%ROOT%requirements.txt" >>"%LOG_DEP%" 2>&1
    if errorlevel 1 (
        echo [ERRO] Falha ao instalar dependencias. Consulte: %LOG_DEP%
        goto FIM_ERRO
    )
)

echo Iniciando motor em modo producao...
echo.
%PYTHON_EXE% "%ROOT%motor_arbitramento.py"
set "RESULT=%ERRORLEVEL%"

if not "%RESULT%"=="0" (
    echo.
    echo ============================================================
    echo EXECUCAO BLOQUEADA OU FINALIZADA COM ERRO. CODIGO %RESULT%
    echo ============================================================
    echo Consulte output e logs_integridade\producao.
    goto FIM_ERRO
)

echo.
echo ============================================================
echo PROCESSAMENTO CONCLUIDO COM SUCESSO
 echo ============================================================
if exist "%ROOT%output\ultimo_arquivo_gerado.txt" (
    set /p "ULTIMO_XLSX="<"%ROOT%output\ultimo_arquivo_gerado.txt"
    echo Ultimo arquivo: !ULTIMO_XLSX!
    if exist "!ULTIMO_XLSX!" start "" "!ULTIMO_XLSX!"
)
exit /b 0

:FIM_ERRO
echo.
echo Pressione qualquer tecla para fechar.
pause >nul
exit /b 1
