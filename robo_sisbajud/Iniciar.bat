@echo off
setlocal DisableDelayedExpansion
set "PYTHONHOME="
set "PYTHONPATH="
set "PYTHONUSERBASE="
set "__PYVENV_LAUNCHER__="
set "PATH=%SystemRoot%\System32;%SystemRoot%\System32\Wbem"
if not exist "%~dp0python\pythonw.exe" goto runtime_ausente
pushd "%~dp0" >nul 2>&1
if errorlevel 1 goto pasta_indisponivel
"%~dp0python\pythonw.exe" -I -B "%~dp0iniciar_portatil.py"
set "ROBO_STATUS=%ERRORLEVEL%"
popd
if not "%ROBO_STATUS%"=="0" goto erro_execucao
exit /b 0
:runtime_ausente
echo O Python portatil nao foi encontrado. Extraia o ZIP inteiro antes de iniciar.
echo O arquivo esperado fica na subpasta python, junto deste iniciador.
pause
exit /b 1
:pasta_indisponivel
echo Nao foi possivel acessar a pasta do robo. Copie a pasta inteira para um local acessivel.
pause
exit /b 1
:erro_execucao
echo O robo nao iniciou corretamente. Consulte a pasta logs e a mensagem exibida.
pause
exit /b %ROBO_STATUS%
