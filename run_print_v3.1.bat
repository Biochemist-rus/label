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
    echo [ERROR] EXE не найден: %EXE%
    exit /b 1
)

:menu
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set TS=%%i
set LOGFILE=%LOGDIR%\print_%TS%.log

cls
echo ==========================================
echo         ПЕЧАТЬ ЭТИКЕТОК v3.1
echo ==========================================
echo EXE:   %EXE%
echo Excel: %EXCEL%
echo Лог:   %LOGFILE%
echo.
echo Выбери принтер:
echo   1 - %PRINTER1%   ^(60x30 мм^) Склад Перерва 11с29
echo   2 - %PRINTER2%   ^(60x40 мм^) Офис Ленинская слобода 19С6
echo.
echo Действия:
echo   3 - Печать одной этикетки
echo   4 - Печать всех этикеток
echo   5 - Отладка одной без печати
echo   6 - Help по командам
echo   7 - Показать текущие настройки
echo   8 - Изменить путь к Excel
echo   9 - Изменить названия колонок
echo   0 - Выход
echo.
echo Текущий принтер: %IP%
set /p CHOICE=Введи пункт: 

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

echo [ERROR] Неизвестный пункт меню
echo [ERROR] Неизвестный пункт меню>>"%LOGFILE%"
exit /b 1

:needip
if not defined IP (
    echo [ERROR] Сначала выбери принтер пунктом 1 или 2
    echo [ERROR] Сначала выбери принтер пунктом 1 или 2>>"%LOGFILE%"
    exit /b 1
)
goto :eof

:baseargs
if "%IP%"=="192.168.14.105" (
    set BASE_ARGS=--excel "%EXCEL%" --column %COLUMN% --gtin-column %GTIN_COLUMN% --name-column %NAME_COLUMN% --ip %IP% --x 15 --y 20 --module-size 3 --text-offset-y 110 --text-width-px 450 --gtin-font-size 22 --name-font-size 24 --ai21-font-size 20 --line-gap 4 --name-max-lines 2 --font-path "C:\Windows\Fonts\arial.ttf" --eac-logo-path eac.png --eac-height-px 64 --eac-gap-px 16 --eac-offset-x -1 --eac-offset-y 0
    goto :eof
)

if "%IP%"=="172.16.102.80" (
    set BASE_ARGS=--excel "%EXCEL%" --column %COLUMN% --gtin-column %GTIN_COLUMN% --name-column %NAME_COLUMN% --ip %IP% --x 80 --y 40 --module-size 4 --text-offset-y 150 --text-width-px 450 --gtin-font-size 22 --name-font-size 26 --ai21-font-size 22 --line-gap 5 --name-max-lines 2 --font-path "C:\Windows\Fonts\arial.ttf" --eac-logo-path eac.png --eac-height-px 64 --eac-gap-px 16 --eac-offset-x 130 --eac-offset-y 0
    goto :eof
)

echo [ERROR] Неизвестный IP принтера: %IP%
echo [ERROR] Неизвестный IP принтера: %IP%>>"%LOGFILE%"
exit /b 1

:one
call :needip
if errorlevel 1 exit /b 1

set /p ROWNUM=Номер строки (0 = первая после заголовка): 
if "%ROWNUM%"=="" set ROWNUM=0

call :baseargs
if errorlevel 1 exit /b 1

echo.>>"%LOGFILE%"
echo ==========================================>>"%LOGFILE%"
echo [RUN] one row=%ROWNUM% ip=%IP% time=%date% %time%>>"%LOGFILE%"
echo COMMAND: %EXE% %BASE_ARGS% --row %ROWNUM%>>"%LOGFILE%"

echo [INFO] Печать одной этикетки...
"%EXE%" %BASE_ARGS% --row %ROWNUM% >>"%LOGFILE%" 2>&1

if errorlevel 1 (
    echo [ERROR] Ошибка печати. Смотри лог:
    echo %LOGFILE%
    echo [ERROR] Ошибка печати>>"%LOGFILE%"
    exit /b 1
)

echo [OK] Готово
pause
goto menu

:all
call :needip
if errorlevel 1 exit /b 1

call :baseargs
if errorlevel 1 exit /b 1

echo.>>"%LOGFILE%"
echo ==========================================>>"%LOGFILE%"
echo [RUN] all ip=%IP% time=%date% %time%>>"%LOGFILE%"
echo COMMAND: %EXE% %BASE_ARGS% --all>>"%LOGFILE%"

echo [INFO] Печать всех этикеток...
"%EXE%" %BASE_ARGS% --all >>"%LOGFILE%" 2>&1

if errorlevel 1 (
    echo [ERROR] Ошибка печати. Смотри лог:
    echo %LOGFILE%
    echo [ERROR] Ошибка печати>>"%LOGFILE%"
    exit /b 1
)

echo [OK] Готово
pause
goto menu

:debugone
call :needip
if errorlevel 1 exit /b 1

set /p ROWNUM=Номер строки для отладки (0 = первая после заголовка): 
if "%ROWNUM%"=="" set ROWNUM=0

call :baseargs
if errorlevel 1 exit /b 1

echo.>>"%LOGFILE%"
echo ==========================================>>"%LOGFILE%"
echo [RUN] debug row=%ROWNUM% ip=%IP% time=%date% %time%>>"%LOGFILE%"
echo COMMAND: %EXE% %BASE_ARGS% --row %ROWNUM% --no-send>>"%LOGFILE%"

echo [INFO] Отладка без печати...
"%EXE%" %BASE_ARGS% --row %ROWNUM% --no-send >>"%LOGFILE%" 2>&1

if errorlevel 1 (
    echo [ERROR] Ошибка отладки. Смотри лог:
    echo %LOGFILE%
    echo [ERROR] Ошибка отладки>>"%LOGFILE%"
    exit /b 1
)

echo [OK] Готово
pause
goto menu

:help
cls
echo ==========================================
echo HELP ПО КОМАНДАМ
echo ==========================================
echo.
echo Профили принтеров:
echo   %PRINTER1% - этикетка 60x30 мм
echo   %PRINTER2% - этикетка 60x40 мм
echo.
echo Обязательные параметры EXE:
echo   --excel "путь_к_excel"
echo   --column barcode
echo   --gtin-column gtin
echo   --name-column name
echo   --ip 192.168.14.105
echo.
echo Режимы:
echo   --row N      печать одной строки
echo   --all        печать всех строк
echo   --no-send    только отладка, без печати
echo.
echo Пример одной строки для 60x30:
echo   %EXE% --excel "%EXCEL%" --column %COLUMN% --gtin-column %GTIN_COLUMN% --name-column %NAME_COLUMN% --row 0 --ip %PRINTER1%
echo.
echo Пример одной строки для 60x40:
echo   %EXE% --excel "%EXCEL%" --column %COLUMN% --gtin-column %GTIN_COLUMN% --name-column %NAME_COLUMN% --row 0 --ip %PRINTER2%
echo.
pause
goto menu

:showsettings
cls
echo ==========================================
echo ТЕКУЩИЕ НАСТРОЙКИ
echo ==========================================
echo EXE          = %EXE%
echo EXCEL        = %EXCEL%
echo COLUMN       = %COLUMN%
echo GTIN_COLUMN  = %GTIN_COLUMN%
echo NAME_COLUMN  = %NAME_COLUMN%
echo PRINTER1     = %PRINTER1%
echo PRINTER2     = %PRINTER2%
echo CURRENT IP   = %IP%
echo LOGFILE      = %LOGFILE%
echo.
pause
goto menu

:setexcel
cls
set /p NEWEXCEL=Новый путь к Excel: 
if not "%NEWEXCEL%"=="" set EXCEL=%NEWEXCEL%
goto menu

:setcolumns
cls
set /p NEWCOL=Колонка Data Matrix [%COLUMN%]: 
if not "%NEWCOL%"=="" set COLUMN=%NEWCOL%
set /p NEWGTIN=Колонка GTIN [%GTIN_COLUMN%]: 
if not "%NEWGTIN%"=="" set GTIN_COLUMN=%NEWGTIN%
set /p NEWNAME=Колонка Название/Артикул [%NAME_COLUMN%]: 
if not "%NEWNAME%"=="" set NAME_COLUMN=%NEWNAME%
goto menu