@echo off
python "%~dp0gerador_xml.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Erro ao executar. Verifique se Python esta instalado.
    pause
)
