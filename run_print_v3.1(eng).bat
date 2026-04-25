@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

cd /d C:\labels

set EXE=C:\labels\dist\print_datamatrix_zpl_graphics_from_excel_v3.1.exe
set EXCEL=C:\labels\data.xlsx
set COLUMN=barcode
set GTIN_COLUMN=gtin
set NAME_COLUMN=name

set PRINTER1=192.168.14.105
set PRINTER2=172.16.102.80

set LOGDIR=C:\labels\logs
if not exist "%LOGDIR%" mkdir "%LOGDIR%"

if not exist "%EXE%" (
    echo [ERROR] EXE not found: %EXE%
    exit /b 1
)

:menu
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set TS=%%i
set LOGFILE=%LOGDIR%\print_%TS%.log

cls
echo ==========================================
echo         LABEL PRINTING TOOL v3.1
echo ==========================================
echo EXE:   %EXE%
echo Excel: %EXCEL%
echo Log:   %LOGFILE%
echo.
echo Select printer:
echo   1 - %PRINTER1%   ^(60x30 mm^)
echo   2 - %PRINTER2%   ^(60x40 mm^)
echo.
echo Actions:
echo   3 - Print ONE label
echo   4 - Print ALL labels
echo   5 - Debug ONE label ^(no send^)
echo   6 - Help
echo   7 - Show current settings
echo   8 - Change Excel path
echo   9 - Change column names
echo   0 - Exit
echo.
echo Current printer: %IP%
set /p CHOICE=Select option: 

if "%CHOICE%"=="1" (
    set IP=%PRINTER1%
    goto menu
)
if "%CHOICE%"=="2" (
    set IP=%PRINTER2%
    goto menu
)
if "%CHOICE%"=="3" goto one
if "%CHOICE%"=="4" goto all
if "%CHOICE%"=="5" goto debugone
if "%CHOICE%"=="6" goto help
if "%CHOICE%"=="7" goto showsettings
if "%CHOICE%"=="8" goto setexcel
if "%CHOICE%"=="9" goto setcolumns
if "%CHOICE%"=="0" exit /b 0

echo [ERROR] Unknown option
echo [ERROR] Unknown option>>"%LOGFILE%"
exit /b 1

:needip
if not defined IP (
    echo [ERROR] Please select printer first (option 1 or 2)
    echo [ERROR] Printer not selected>>"%LOGFILE%"
    exit /b 1
)
goto :eof

:baseargs
if "%IP%"=="192.168.14.105" (
    set BASE_ARGS=--excel "%EXCEL%" --column %COLUMN% --gtin-column %GTIN_COLUMN% --name-column %NAME_COLUMN% --ip %IP% --x 32 --y 20 --module-size 3 --text-offset-y 110 --text-width-px 450 --gtin-font-size 22 --name-font-size 28 --ai21-font-size 20 --line-gap 4 --name-max-lines 2 --font-path "C:\Windows\Fonts\arial.ttf" --eac-logo-path eac.png --eac-height-px 64 --eac-gap-px 16 --eac-offset-x -1 --eac-offset-y 0
    goto :eof
)

if "%IP%"=="172.16.102.80" (
    set BASE_ARGS=--excel "%EXCEL%" --column %COLUMN% --gtin-column %GTIN_COLUMN% --name-column %NAME_COLUMN% --ip %IP% --x 100 --y 40 --module-size 4 --text-offset-y 150 --text-width-px 450 --gtin-font-size 24 --name-font-size 30 --ai21-font-size 22 --line-gap 5 --name-max-lines 2 --font-path "C:\Windows\Fonts\arial.ttf" --eac-logo-path eac.png --eac-height-px 64 --eac-gap-px 16 --eac-offset-x 130 --eac-offset-y 0
    goto :eof
)

echo [ERROR] Unknown printer IP: %IP%
exit /b 1

:run_and_show
set RUN_ARGS=%~1

del /q "%TEMP%\print_stdout.txt" 2>nul
del /q "%TEMP%\print_stderr.txt" 2>nul

powershell -NoProfile -Command ^
  "$p = Start-Process -FilePath '%EXE%' -ArgumentList '%RUN_ARGS%' -NoNewWindow -RedirectStandardOutput '%TEMP%\print_stdout.txt' -RedirectStandardError '%TEMP%\print_stderr.txt' -PassThru -Wait; exit $p.ExitCode"

set EXITCODE=%ERRORLEVEL%

if exist "%TEMP%\print_stdout.txt" (
    type "%TEMP%\print_stdout.txt"
    type "%TEMP%\print_stdout.txt" >>"%LOGFILE%"
)

if exist "%TEMP%\print_stderr.txt" (
    type "%TEMP%\print_stderr.txt"
    type "%TEMP%\print_stderr.txt" >>"%LOGFILE%"
)

exit /b %EXITCODE%

:one
call :needip
if errorlevel 1 exit /b 1

set /p ROWNUM=Enter row number (0 = first data row): 
if "%ROWNUM%"=="" set ROWNUM=0

call :baseargs
if errorlevel 1 exit /b 1

echo [INFO] Printing one label...
call :run_and_show "%BASE_ARGS% --row %ROWNUM%"

if errorlevel 1 (
    echo [ERROR] Print failed. See log:
    echo %LOGFILE%
    exit /b 1
)

echo [OK] Done
pause
goto menu

:all
call :needip
if errorlevel 1 exit /b 1

call :baseargs
if errorlevel 1 exit /b 1

echo [INFO] Printing all labels...
call :run_and_show "%BASE_ARGS% --all"

if errorlevel 1 (
    echo [ERROR] Print failed. See log:
    echo %LOGFILE%
    exit /b 1
)

echo [OK] Done
pause
goto menu

:debugone
call :needip
if errorlevel 1 exit /b 1

set /p ROWNUM=Enter row for debug: 
if "%ROWNUM%"=="" set ROWNUM=0

call :baseargs
if errorlevel 1 exit /b 1

echo [INFO] Debug (no send)...
call :run_and_show "%BASE_ARGS% --row %ROWNUM% --no-send"

if errorlevel 1 (
    echo [ERROR] Debug failed. See log:
    echo %LOGFILE%
    exit /b 1
)

echo [OK] Done
pause
goto menu

:help
cls
echo ==========================================
echo HELP
echo ==========================================
echo.
echo This tool prints labels from Excel using ZPL.
echo.
echo Required EXE parameters:
echo   --excel "path_to_excel"
echo   --column barcode
echo   --gtin-column gtin
echo   --name-column name
echo   --ip printer_ip
echo.
echo Modes:
echo   --row N      print single row
echo   --all        print all rows
echo   --no-send    debug only (no printing)
echo.
echo Printers:
echo   %PRINTER1% - label 60x30 mm
echo   %PRINTER2% - label 60x40 mm
echo.
pause
goto menu

:showsettings
cls
echo ==========================================
echo CURRENT SETTINGS
echo ==========================================
echo EXE          = %EXE%
echo EXCEL        = %EXCEL%
echo COLUMN       = %COLUMN%
echo GTIN_COLUMN  = %GTIN_COLUMN%
echo NAME_COLUMN  = %NAME_COLUMN%
echo CURRENT IP   = %IP%
echo LOGFILE      = %LOGFILE%
echo.
pause
goto menu

:setexcel
set /p NEWEXCEL=Enter new Excel path: 
if not "%NEWEXCEL%"=="" set EXCEL=%NEWEXCEL%
goto menu

:setcolumns
set /p NEWCOL=Barcode column [%COLUMN%]: 
if not "%NEWCOL%"=="" set COLUMN=%NEWCOL%
set /p NEWGTIN=GTIN column [%GTIN_COLUMN%]: 
if not "%NEWGTIN%"=="" set GTIN_COLUMN=%NEWGTIN%
set /p NEWNAME=Name column [%NAME_COLUMN%]: 
if not "%NEWNAME%"=="" set NAME_COLUMN=%NEWNAME%
goto menu