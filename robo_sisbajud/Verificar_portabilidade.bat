@echo off
setlocal DisableDelayedExpansion
set "PYTHONHOME="
set "PYTHONPATH="
set "PYTHONUSERBASE="
set "__PYVENV_LAUNCHER__="
set "PATH=%SystemRoot%\System32;%SystemRoot%\System32\Wbem"
if not exist "%~dp0python\python.exe" goto runtime_ausente
pushd "%~dp0" >nul 2>&1
if errorlevel 1 goto pasta_indisponivel
if /I "%~1"=="--clipboard" goto com_clipboard
"%~dp0python\python.exe" -I -B "%~dp0verificar_portabilidade.py"
goto resultado
:com_clipboard
"%~dp0python\python.exe" -I -B "%~dp0verificar_portabilidade.py" --clipboard
:resultado
set "ROBO_STATUS=%ERRORLEVEL%"
popd
echo.
echo Consulte logs\relatorio_portabilidade.json para os detalhes.
pause
exit /b %ROBO_STATUS%
:runtime_ausente
echo Python portatil ausente. Extraia o ZIP inteiro antes de verificar.
pause
exit /b 1
:pasta_indisponivel
echo Nao foi possivel acessar a pasta do robo.
pause
exit /b 1
