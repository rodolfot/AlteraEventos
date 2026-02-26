@echo off
setlocal

set MAVEN_VERSION=3.9.6
set MAVEN_DIR=%~dp0.maven
set MVN_CMD=%MAVEN_DIR%\apache-maven-%MAVEN_VERSION%\bin\mvn.cmd
set MAVEN_URL=https://downloads.apache.org/maven/maven-3/%MAVEN_VERSION%/binaries/apache-maven-%MAVEN_VERSION%-bin.zip
set MAVEN_ZIP=%MAVEN_DIR%\maven.zip

echo ===========================================
echo  AlteraEventos - Build Script
echo ===========================================

:: Verifica se mvn ja esta disponivel no PATH
where mvn >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [OK] Maven encontrado no PATH.
    set MVN_CMD=mvn
    goto BUILD
)

:: Verifica se ja foi baixado localmente
if exist "%MVN_CMD%" (
    echo [OK] Maven local encontrado em: %MAVEN_DIR%
    goto BUILD
)

:: Baixa Maven
echo [INFO] Maven nao encontrado. Baixando Apache Maven %MAVEN_VERSION%...
if not exist "%MAVEN_DIR%" mkdir "%MAVEN_DIR%"

powershell -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%MAVEN_URL%' -OutFile '%MAVEN_ZIP%' -UseBasicParsing }"

if not exist "%MAVEN_ZIP%" (
    echo [ERRO] Falha ao baixar Maven. Verifique sua conexao com a internet.
    echo Baixe manualmente em: https://maven.apache.org/download.cgi
    pause
    exit /b 1
)

echo [INFO] Extraindo Maven...
powershell -Command "Expand-Archive -Path '%MAVEN_ZIP%' -DestinationPath '%MAVEN_DIR%' -Force"
del "%MAVEN_ZIP%"

if not exist "%MVN_CMD%" (
    echo [ERRO] Falha ao extrair Maven.
    pause
    exit /b 1
)

echo [OK] Maven instalado em: %MAVEN_DIR%

:BUILD
echo.
echo [INFO] Compilando e empacotando...
echo.
"%MVN_CMD%" clean package -q

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERRO] Build falhou. Execute com detalhes:
    echo   build.bat clean package  (sem -q para ver erros)
    pause
    exit /b 1
)

echo.
echo ===========================================
echo  Build concluido com sucesso!
echo  JAR: target\AlteraEventos.jar
echo ===========================================
echo.
echo Para executar:
echo   java -jar target\AlteraEventos.jar
echo.

set /p EXECUTAR="Deseja executar agora? (S/N): "
if /i "%EXECUTAR%"=="S" (
    java -jar "%~dp0target\AlteraEventos.jar"
)

endlocal
