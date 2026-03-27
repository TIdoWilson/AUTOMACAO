@echo off
setlocal EnableExtensions

title Coleta de Informacoes do PC

REM Defina aqui o caminho fixo onde os .txt serao salvos:
set "DESTINO_FIXO=C:\RelatoriosHardware"

echo.
set /p "NOME=Digite o nome do arquivo (sem .txt): "

if not defined NOME (
    echo Nome invalido. Encerrando.
    exit /b 1
)

echo(%NOME%| findstr /r "[\\/:*?\"<>|]" >nul
if not errorlevel 1 (
    echo O nome contem caracteres invalidos para arquivo. Encerrando.
    exit /b 1
)

if not exist "%DESTINO_FIXO%" mkdir "%DESTINO_FIXO%"

set "ARQUIVO=%DESTINO_FIXO%\%NOME%.txt"

(
    echo ========================================
    echo RELATORIO DE HARDWARE DO PC
    echo ========================================
    echo Data/Hora: %date% %time%
    echo Computador: %COMPUTERNAME%
    echo Usuario: %USERNAME%
    echo.

    echo [PROCESSADOR]
    wmic cpu get Name,Manufacturer,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed /value
    echo.

    echo [PLACA-MAE]
    wmic baseboard get Manufacturer,Product,SerialNumber,Version /value
    echo.

    echo [MEMORIA RAM - TOTAL]
    wmic computersystem get TotalPhysicalMemory /value
    echo.

    echo [MEMORIA RAM - MODULOS]
    wmic memorychip get BankLabel,Manufacturer,PartNumber,Capacity,Speed,SerialNumber /value
) > "%ARQUIVO%"

if errorlevel 1 (
    echo Falha ao gerar o arquivo.
    exit /b 1
)

echo.
echo Arquivo gerado com sucesso:
echo %ARQUIVO%

echo.
pause
