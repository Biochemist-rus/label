@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

cd /d C:\labels

set EXE=C:\labels\dist\print_datamatrix_zpl_graphics_from_excel_v3.1.exe
set INI=C:\labels\menu_print_v3.1.ini
set LOGDIR=C:\labels\logs

if not exist "%LOGDIR%" mkdir "%LOGDIR%"

REM === Default fallback values ===
set PRINTER1=192.168.14.105
set PRINTER2=172.16.102.80
set EXCEL=C:\labels\data.xlsx
set COLUMN=barcode
set GTIN_COLUMN=gtin
set NAME_COLUMN=name
set LAST_ROW=0

if not exist "%EXE%" (
    echo [ERROR] EXE not found: %EXE%
    exit /b 1
)

if not exist "%INI%" call :create_default_ini
call :load_ini

:menu
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set TS=%%i
set LOGFILE=%LOGDIR%\print_%TS%.log

cls
echo ==========================================
echo         LABEL PRINTING TOOL v3.1
echo ==========================================
echo EXE:   %EXE%
echo INI:   %INI%
echo Log:   %LOGFILE%
echo.
echo Current printer:
echo   %IP%
echo.
echo Printers:
echo   1 - %PRINTER1%
echo   2 - %PRINTER2%
echo.
echo Settings:
echo   Excel  = %EXCEL%
echo   Columns= %COLUMN% / %GTIN_COLUMN% / %NAME_COLUMN%
echo   Last row = %LAST_ROW%
echo.
echo Actions:
echo   3 - Print ONE
echo   4 - Print ALL
echo   5 - Debug ONE
echo   6 - Print LAST
echo   7 - Help
echo   8 - Show settings
echo   9 - Change Excel
echo   A - Change columns
echo   P - Change printer IPs
echo   S - Save
echo   R - Reload
echo   0 - Exit
echo.

set /p CHOICE=Select option: 

if /I "%CHOICE%"=="1" (set IP=%PRINTER1% & call :save_ini & goto menu)
if /I "%CHOICE%"=="2" (set IP=%PRINTER2% & call :save_ini & goto menu)
if /I "%CHOICE%"=="3" goto one
if /I "%CHOICE%"=="4" goto all
if /I "%CHOICE%"=="5" goto debugone
if /I "%CHOICE%"=="6" goto lastrow
if /I "%CHOICE%"=="7" goto help
if /I "%CHOICE%"=="8" goto show
if /I "%CHOICE%"=="9" goto setexcel
if /I "%CHOICE%"=="A" goto setcolumns
if /I "%CHOICE%"=="P" goto setprinters
if /I "%CHOICE%"=="S" (call :save_ini & echo Saved & pause & goto menu)
if /I "%CHOICE%"=="R" (call :load_ini & echo Reloaded & pause & goto menu)
if /I "%CHOICE%"=="0" exit /b 0

echo [ERROR] Unknown option
exit /b 1

:baseargs
if "%IP%"=="%PRINTER1%" (
    set BASE_ARGS=--excel "%EXCEL%" --column %COLUMN% --gtin-column %GTIN_COLUMN% --name-column %NAME_COLUMN% --ip %IP% --x 15 --y 20 --module-size 3 --text-offset-y 110 --text-width-px 450 --gtin-font-size 22 --name-font-size 24 --ai21-font-size 20 --line-gap 4 --name-max-lines 2 --font-path "C:\Windows\Fonts\arial.ttf" --eac-logo-path eac.png --eac-height-px 64 --eac-gap-px 16 --eac-offset-x -1 --eac-offset-y 0
    goto :eof
)

if "%IP%"=="%PRINTER2%" (
    set BASE_ARGS=--excel "%EXCEL%" --column %COLUMN% --gtin-column %GTIN_COLUMN% --name-column %NAME_COLUMN% --ip %IP% --x 80 --y 40 --module-size 4 --text-offset-y 150 --text-width-px 450 --gtin-font-size 22 --name-font-size 26 --ai21-font-size 22 --line-gap 5 --name-max-lines 2 --font-path "C:\Windows\Fonts\arial.ttf" --eac-logo-path eac.png --eac-height-px 64 --eac-gap-px 16 --eac-offset-x 130 --eac-offset-y 0
    goto :eof
)

echo [ERROR] Unknown IP
exit /b 1

:run
set ARGS=%~1

powershell -NoProfile -Command ^
 "$p=Start-Process '%EXE%' '%ARGS%' -NoNewWindow -PassThru -Wait; exit $p.ExitCode"

exit /b %ERRORLEVEL%

:one
set /p ROWNUM=Row: 
if "%ROWNUM%"=="" set ROWNUM=0
set LAST_ROW=%ROWNUM%
call :save_ini
call :baseargs
call :run "%BASE_ARGS% --row %ROWNUM%"
if errorlevel 1 exit /b 1
pause
goto menu

:lastrow
call :baseargs
call :run "%BASE_ARGS% --row %LAST_ROW%"
if errorlevel 1 exit /b 1
pause
goto menu

:all
call :baseargs
call :run "%BASE_ARGS% --all"
if errorlevel 1 exit /b 1
pause
goto menu

:debugone
set /p ROWNUM=Row: 
if "%ROWNUM%"=="" set ROWNUM=0
set LAST_ROW=%ROWNUM%
call :save_ini
call :baseargs
call :run "%BASE_ARGS% --row %ROWNUM% --no-send"
if errorlevel 1 exit /b 1
pause
goto menu

:setexcel
set /p EXCEL=New Excel: 
goto menu

:setcolumns
set /p COLUMN=Barcode: 
set /p GTIN_COLUMN=GTIN: 
set /p NAME_COLUMN=Name: 
goto menu

:setprinters
set /p PRINTER1=Printer1 IP: 
set /p PRINTER2=Printer2 IP: 
goto menu

:show
echo Excel=%EXCEL%
echo Columns=%COLUMN% %GTIN_COLUMN% %NAME_COLUMN%
echo Printers=%PRINTER1% %PRINTER2%
echo Current=%IP%
echo LastRow=%LAST_ROW%
pause
goto menu

:help
echo Uses EXE to print labels from Excel
echo Modes: --row --all --no-send
pause
goto menu

:create_default_ini
(
echo [Settings]
echo ExcelFile=C:\labels\data.xlsx
echo BarcodeColumn=barcode
echo GtinColumn=gtin
echo NameColumn=name
echo Printer1IP=192.168.14.105
echo Printer2IP=172.16.102.80
echo CurrentIP=192.168.14.105
echo LastRow=0
)> "%INI%"
goto :eof

:load_ini
for /f "tokens=1,2 delims==" %%A in (%INI%) do (
if "%%A"=="ExcelFile" set EXCEL=%%B
if "%%A"=="BarcodeColumn" set COLUMN=%%B
if "%%A"=="GtinColumn" set GTIN_COLUMN=%%B
if "%%A"=="NameColumn" set NAME_COLUMN=%%B
if "%%A"=="Printer1IP" set PRINTER1=%%B
if "%%A"=="Printer2IP" set PRINTER2=%%B
if "%%A"=="CurrentIP" set IP=%%B
if "%%A"=="LastRow" set LAST_ROW=%%B
)
goto :eof

:save_ini
(
echo [Settings]
echo ExcelFile=%EXCEL%
echo BarcodeColumn=%COLUMN%
echo GtinColumn=%GTIN_COLUMN%
echo NameColumn=%NAME_COLUMN%
echo Printer1IP=%PRINTER1%
echo Printer2IP=%PRINTER2%
echo CurrentIP=%IP%
echo LastRow=%LAST_ROW%
)> "%INI%"
goto :eof