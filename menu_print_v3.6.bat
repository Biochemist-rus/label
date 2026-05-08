@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

cd /d "%~dp0"
set "ROOT=%CD%"

set "VERSION=3.6"
set "EXE=dist\print_datamatrix_zpl_graphics_from_excel_v3.6.exe"
set "INI=menu_print_v3.6.ini"
set "LOGDIR=logs"

if not exist "%LOGDIR%" mkdir "%LOGDIR%"

set "LANG=EN"
set "EXCEL=data.xlsx"
set "COLUMN=barcode"
set "GTIN_COLUMN=gtin"
set "NAME_COLUMN=name"
set "PRINTER1=172.16.104.99"
set "PRINTER2=192.168.14.105"
set "IP=172.16.104.99"
set "LAST_ROW=0"

set "P1_LABEL=60x30"
set "P1_X=15"
set "P1_Y=20"
set "P1_MODULE_SIZE=3"
set "P1_TEXT_OFFSET_Y=110"
set "P1_TEXT_WIDTH_PX=450"
set "P1_GTIN_FONT_SIZE=20"
set "P1_NAME_FONT_SIZE=22"
set "P1_AI21_FONT_SIZE=20"
set "P1_LINE_GAP=4"
set "P1_NAME_MAX_LINES=2"
set "P1_FONT_PATH=fonts\arial.ttf"
set "P1_EAC_LOGO_PATH=eac.png"
set "P1_EAC_HEIGHT_PX=64"
set "P1_EAC_GAP_PX=50"
set "P1_EAC_OFFSET_X=95"
set "P1_EAC_OFFSET_Y=0"

set "P2_LABEL=60x40"
set "P2_X=80"
set "P2_Y=40"
set "P2_MODULE_SIZE=4"
set "P2_TEXT_OFFSET_Y=150"
set "P2_TEXT_WIDTH_PX=450"
set "P2_GTIN_FONT_SIZE=22"
set "P2_NAME_FONT_SIZE=26"
set "P2_AI21_FONT_SIZE=22"
set "P2_LINE_GAP=5"
set "P2_NAME_MAX_LINES=2"
set "P2_FONT_PATH=fonts\arial.ttf"
set "P2_EAC_LOGO_PATH=eac.png"
set "P2_EAC_HEIGHT_PX=64"
set "P2_EAC_GAP_PX=16"
set "P2_EAC_OFFSET_X=130"
set "P2_EAC_OFFSET_Y=0"

if not exist "%INI%" call :save_ini
call :load_ini

if not exist "%EXE%" (
  echo [ERROR] EXE not found: %ROOT%\%EXE%
  echo [ОШИБКА] EXE не найден: %ROOT%\%EXE%
  echo.
  echo Build it first with: build_exe_v3.6.bat
  echo Сначала соберите его через: build_exe_v3.6.bat
  pause
  exit /b 1
)

call :choose_language

:menu
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "TS=%%i"
set "LOGFILE=%LOGDIR%\print_%TS%.log"
cls
if /I "%LANG%"=="RU" goto menu_ru

:menu_en
echo ==========================================
echo         LABEL PRINTING TOOL v%VERSION%
echo ==========================================
echo Root:  %ROOT%
echo EXE:   %EXE%
echo INI:   %INI%
echo Log:   %LOGFILE%
echo.
echo Current printer: %IP%
echo   1 - %PRINTER1% [%P1_LABEL%]
echo   2 - %PRINTER2% [%P2_LABEL%]
echo.
echo Excel:   %EXCEL%
echo Columns: %COLUMN% / %GTIN_COLUMN% / %NAME_COLUMN%
echo Last row: %LAST_ROW%
echo.
echo 3 - Print ONE
echo 4 - Print ALL
echo 5 - Debug ONE
echo 6 - Print LAST
echo 7 - Help
echo 8 - Show settings
echo 9 - Change Excel
echo A - Change columns
echo P - Change printer IPs
echo L - Change language
echo S - Save INI
echo R - Reload INI
echo 0 - Exit
echo.
set /p "CHOICE=Select option: "
goto handle_choice

:menu_ru
echo ==========================================
echo      ПЕЧАТЬ ЭТИКЕТОК v%VERSION%
echo ==========================================
echo Корень: %ROOT%
echo EXE:    %EXE%
echo INI:    %INI%
echo Лог:    %LOGFILE%
echo.
echo Текущий принтер: %IP%
echo   1 - %PRINTER1% [%P1_LABEL%]
echo   2 - %PRINTER2% [%P2_LABEL%]
echo.
echo Excel:  %EXCEL%
echo Колонки: %COLUMN% / %GTIN_COLUMN% / %NAME_COLUMN%
echo Последняя строка: %LAST_ROW%
echo.
echo 3 - Печать одной этикетки
echo 4 - Печать всех строк
echo 5 - Отладка одной строки без печати
echo 6 - Печать последней строки
echo 7 - Помощь
echo 8 - Показать настройки
echo 9 - Изменить Excel-файл
echo A - Изменить имена колонок
echo P - Изменить IP принтеров
echo L - Сменить язык
echo S - Сохранить INI
echo R - Перезагрузить INI
echo 0 - Выход
echo.
set /p "CHOICE=Выберите пункт: "
goto handle_choice

:handle_choice
if /I "%CHOICE%"=="1" (set "IP=%PRINTER1%" & call :save_ini & goto menu)
if /I "%CHOICE%"=="2" (set "IP=%PRINTER2%" & call :save_ini & goto menu)
if /I "%CHOICE%"=="3" goto one
if /I "%CHOICE%"=="4" goto all
if /I "%CHOICE%"=="5" goto debugone
if /I "%CHOICE%"=="6" goto lastrow
if /I "%CHOICE%"=="7" goto help
if /I "%CHOICE%"=="8" goto show
if /I "%CHOICE%"=="9" goto setexcel
if /I "%CHOICE%"=="A" goto setcolumns
if /I "%CHOICE%"=="P" goto setprinters
if /I "%CHOICE%"=="L" goto langmenu
if /I "%CHOICE%"=="S" (call :save_ini & call :msg_saved & pause & goto menu)
if /I "%CHOICE%"=="R" (call :load_ini & call :msg_reloaded & pause & goto menu)
if /I "%CHOICE%"=="0" exit /b 0
call :msg_unknown_option
pause
goto menu

:choose_language
cls
echo ==========================================
echo Select language / Выберите язык
echo ==========================================
echo 1 - English
echo 2 - Русский
echo.
set /p "LANG_CHOICE=Language / Язык [1/2, Enter=%LANG%]: "
if "%LANG_CHOICE%"=="" goto :eof
if "%LANG_CHOICE%"=="1" set "LANG=EN"
if "%LANG_CHOICE%"=="2" set "LANG=RU"
call :save_ini
goto :eof

:langmenu
call :choose_language
goto menu

:select_profile
if "%IP%"=="%PRINTER1%" (
  set "P_LABEL=%P1_LABEL%"
  set "P_X=%P1_X%"
  set "P_Y=%P1_Y%"
  set "P_MODULE_SIZE=%P1_MODULE_SIZE%"
  set "P_TEXT_OFFSET_Y=%P1_TEXT_OFFSET_Y%"
  set "P_TEXT_WIDTH_PX=%P1_TEXT_WIDTH_PX%"
  set "P_GTIN_FONT_SIZE=%P1_GTIN_FONT_SIZE%"
  set "P_NAME_FONT_SIZE=%P1_NAME_FONT_SIZE%"
  set "P_AI21_FONT_SIZE=%P1_AI21_FONT_SIZE%"
  set "P_LINE_GAP=%P1_LINE_GAP%"
  set "P_NAME_MAX_LINES=%P1_NAME_MAX_LINES%"
  set "P_FONT_PATH=%P1_FONT_PATH%"
  set "P_EAC_LOGO_PATH=%P1_EAC_LOGO_PATH%"
  set "P_EAC_HEIGHT_PX=%P1_EAC_HEIGHT_PX%"
  set "P_EAC_GAP_PX=%P1_EAC_GAP_PX%"
  set "P_EAC_OFFSET_X=%P1_EAC_OFFSET_X%"
  set "P_EAC_OFFSET_Y=%P1_EAC_OFFSET_Y%"
  goto :eof
)
if "%IP%"=="%PRINTER2%" (
  set "P_LABEL=%P2_LABEL%"
  set "P_X=%P2_X%"
  set "P_Y=%P2_Y%"
  set "P_MODULE_SIZE=%P2_MODULE_SIZE%"
  set "P_TEXT_OFFSET_Y=%P2_TEXT_OFFSET_Y%"
  set "P_TEXT_WIDTH_PX=%P2_TEXT_WIDTH_PX%"
  set "P_GTIN_FONT_SIZE=%P2_GTIN_FONT_SIZE%"
  set "P_NAME_FONT_SIZE=%P2_NAME_FONT_SIZE%"
  set "P_AI21_FONT_SIZE=%P2_AI21_FONT_SIZE%"
  set "P_LINE_GAP=%P2_LINE_GAP%"
  set "P_NAME_MAX_LINES=%P2_NAME_MAX_LINES%"
  set "P_FONT_PATH=%P2_FONT_PATH%"
  set "P_EAC_LOGO_PATH=%P2_EAC_LOGO_PATH%"
  set "P_EAC_HEIGHT_PX=%P2_EAC_HEIGHT_PX%"
  set "P_EAC_GAP_PX=%P2_EAC_GAP_PX%"
  set "P_EAC_OFFSET_X=%P2_EAC_OFFSET_X%"
  set "P_EAC_OFFSET_Y=%P2_EAC_OFFSET_Y%"
  goto :eof
)
call :msg_unknown_ip
exit /b 1

:baseargs
call :select_profile
if errorlevel 1 exit /b 1
set "BASE_ARGS=--excel "%EXCEL%" --column "%COLUMN%" --gtin-column "%GTIN_COLUMN%" --name-column "%NAME_COLUMN%" --ip %IP% --x %P_X% --y %P_Y% --module-size %P_MODULE_SIZE% --text-offset-y %P_TEXT_OFFSET_Y% --text-width-px %P_TEXT_WIDTH_PX% --gtin-font-size %P_GTIN_FONT_SIZE% --name-font-size %P_NAME_FONT_SIZE% --ai21-font-size %P_AI21_FONT_SIZE% --line-gap %P_LINE_GAP% --name-max-lines %P_NAME_MAX_LINES% --font-path "%P_FONT_PATH%" --eac-logo-path "%P_EAC_LOGO_PATH%" --eac-height-px %P_EAC_HEIGHT_PX% --eac-gap-px %P_EAC_GAP_PX% --eac-offset-x %P_EAC_OFFSET_X% --eac-offset-y %P_EAC_OFFSET_Y%"
goto :eof

:run
(
  echo ==========================================
  echo [RUN] %DATE% %TIME%
  echo ROOT=%ROOT%
  echo EXE=%EXE%
  echo IP=%IP%
  echo LABEL=%P_LABEL%
  echo ARGS=%ARGS%
  echo ==========================================
) > "%LOGFILE%"
"%EXE%" %ARGS% >> "%LOGFILE%" 2>&1
set "RC=%ERRORLEVEL%"
type "%LOGFILE%"
if not "%RC%"=="0" call :msg_run_failed
exit /b %RC%

:one
if /I "%LANG%"=="RU" (set /p "ROWNUM=Строка: ") else (set /p "ROWNUM=Row: ")
if "%ROWNUM%"=="" set "ROWNUM=0"
set "LAST_ROW=%ROWNUM%"
call :save_ini
call :baseargs
if errorlevel 1 pause & goto menu
set "ARGS=%BASE_ARGS% --row %ROWNUM%"
call :run
pause
goto menu

:lastrow
call :baseargs
if errorlevel 1 pause & goto menu
set "ARGS=%BASE_ARGS% --row %LAST_ROW%"
call :run
pause
goto menu

:all
call :baseargs
if errorlevel 1 pause & goto menu
set "ARGS=%BASE_ARGS% --all"
call :run
pause
goto menu

:debugone
if /I "%LANG%"=="RU" (set /p "ROWNUM=Строка: ") else (set /p "ROWNUM=Row: ")
if "%ROWNUM%"=="" set "ROWNUM=0"
set "LAST_ROW=%ROWNUM%"
call :save_ini
call :baseargs
if errorlevel 1 pause & goto menu
set "ARGS=%BASE_ARGS% --row %ROWNUM% --no-send"
call :run
pause
goto menu

:setexcel
if /I "%LANG%"=="RU" (set /p "EXCEL=Excel-файл относительно папки BAT: ") else (set /p "EXCEL=Excel file relative to BAT folder: ")
call :save_ini
goto menu

:setcolumns
if /I "%LANG%"=="RU" (
  set /p "COLUMN=Колонка Data Matrix: "
  set /p "GTIN_COLUMN=Колонка GTIN: "
  set /p "NAME_COLUMN=Колонка названия: "
) else (
  set /p "COLUMN=Barcode column: "
  set /p "GTIN_COLUMN=GTIN column: "
  set /p "NAME_COLUMN=Name column: "
)
call :save_ini
goto menu

:setprinters
if /I "%LANG%"=="RU" (
  set /p "PRINTER1=IP принтера 1 [%P1_LABEL%]: "
  set /p "PRINTER2=IP принтера 2 [%P2_LABEL%]: "
) else (
  set /p "PRINTER1=Printer 1 IP [%P1_LABEL%]: "
  set /p "PRINTER2=Printer 2 IP [%P2_LABEL%]: "
)
call :save_ini
goto menu

:show
call :select_profile
echo Root=%ROOT%
echo EXE=%EXE%
echo INI=%INI%
echo Excel=%EXCEL%
echo Columns=%COLUMN% / %GTIN_COLUMN% / %NAME_COLUMN%
echo Printer1=%PRINTER1% [%P1_LABEL%]
echo Printer2=%PRINTER2% [%P2_LABEL%]
echo Current=%IP% [%P_LABEL%]
echo LastRow=%LAST_ROW%
echo Language=%LANG%
echo.
echo Current profile values:
echo x=%P_X% y=%P_Y% module-size=%P_MODULE_SIZE% text-offset-y=%P_TEXT_OFFSET_Y%
echo text-width-px=%P_TEXT_WIDTH_PX% gtin-font-size=%P_GTIN_FONT_SIZE% name-font-size=%P_NAME_FONT_SIZE% ai21-font-size=%P_AI21_FONT_SIZE%
echo line-gap=%P_LINE_GAP% name-max-lines=%P_NAME_MAX_LINES%
echo font-path=%P_FONT_PATH%
echo eac-logo-path=%P_EAC_LOGO_PATH% eac-height-px=%P_EAC_HEIGHT_PX% eac-gap-px=%P_EAC_GAP_PX% eac-offset-x=%P_EAC_OFFSET_X% eac-offset-y=%P_EAC_OFFSET_Y%
pause
goto menu

:help
if /I "%LANG%"=="RU" goto help_ru
echo This menu calls the v%VERSION% EXE and prints labels from Excel.
echo All paths are relative to the folder containing this BAT file.
echo Expected layout:
echo   menu_print_v3.6.bat
echo   menu_print_v3.6.ini
echo   data.xlsx
echo   eac.png
echo   fonts\arial.ttf
echo   dist\print_datamatrix_zpl_graphics_from_excel_v3.6.exe
echo Blank barcode rows feed one empty label.
pause
goto menu

:help_ru
echo Это меню запускает EXE v%VERSION% и печатает этикетки из Excel.
echo Все пути относительны папке, где лежит этот BAT-файл.
echo Ожидаемая структура:
echo   menu_print_v3.6.bat
echo   menu_print_v3.6.ini
echo   data.xlsx
echo   eac.png
echo   fonts\arial.ttf
echo   dist\print_datamatrix_zpl_graphics_from_excel_v3.6.exe
echo Пустая строка barcode протягивает одну пустую наклейку.
pause
goto menu

:msg_saved
if /I "%LANG%"=="RU" (echo Сохранено) else (echo Saved)
goto :eof

:msg_reloaded
if /I "%LANG%"=="RU" (echo INI перезагружен) else (echo Reloaded)
goto :eof

:msg_unknown_option
if /I "%LANG%"=="RU" (echo [ОШИБКА] Неизвестный пункт меню) else (echo [ERROR] Unknown option)
goto :eof

:msg_unknown_ip
if /I "%LANG%"=="RU" (echo [ОШИБКА] Неизвестный IP принтера: %IP%) else (echo [ERROR] Unknown printer IP: %IP%)
goto :eof

:msg_run_failed
if /I "%LANG%"=="RU" (echo [ОШИБКА] Печать завершилась с кодом %RC%. См. лог: %LOGFILE%) else (echo [ERROR] Print command failed with code %RC%. See log: %LOGFILE%)
goto :eof

:load_ini
for /f "usebackq tokens=1,* delims==" %%A in ("%INI%") do (
  set "KEY=%%~A"
  set "VAL=%%~B"
  if not "!KEY!"=="" if not "!KEY:~0,1!"==";" if not "!KEY:~0,1!"=="#" if not "!KEY:~0,1!"=="[" (
    if /I "!KEY!"=="Language" set "LANG=!VAL!"
    if /I "!KEY!"=="ExcelFile" set "EXCEL=!VAL!"
    if /I "!KEY!"=="BarcodeColumn" set "COLUMN=!VAL!"
    if /I "!KEY!"=="GtinColumn" set "GTIN_COLUMN=!VAL!"
    if /I "!KEY!"=="NameColumn" set "NAME_COLUMN=!VAL!"
    if /I "!KEY!"=="Printer1IP" set "PRINTER1=!VAL!"
    if /I "!KEY!"=="Printer2IP" set "PRINTER2=!VAL!"
    if /I "!KEY!"=="CurrentIP" set "IP=!VAL!"
    if /I "!KEY!"=="LastRow" set "LAST_ROW=!VAL!"
    if /I "!KEY!"=="Printer1Label" set "P1_LABEL=!VAL!"
    if /I "!KEY!"=="Printer1X" set "P1_X=!VAL!"
    if /I "!KEY!"=="Printer1Y" set "P1_Y=!VAL!"
    if /I "!KEY!"=="Printer1ModuleSize" set "P1_MODULE_SIZE=!VAL!"
    if /I "!KEY!"=="Printer1TextOffsetY" set "P1_TEXT_OFFSET_Y=!VAL!"
    if /I "!KEY!"=="Printer1TextWidthPx" set "P1_TEXT_WIDTH_PX=!VAL!"
    if /I "!KEY!"=="Printer1GtinFontSize" set "P1_GTIN_FONT_SIZE=!VAL!"
    if /I "!KEY!"=="Printer1NameFontSize" set "P1_NAME_FONT_SIZE=!VAL!"
    if /I "!KEY!"=="Printer1Ai21FontSize" set "P1_AI21_FONT_SIZE=!VAL!"
    if /I "!KEY!"=="Printer1LineGap" set "P1_LINE_GAP=!VAL!"
    if /I "!KEY!"=="Printer1NameMaxLines" set "P1_NAME_MAX_LINES=!VAL!"
    if /I "!KEY!"=="Printer1FontPath" set "P1_FONT_PATH=!VAL!"
    if /I "!KEY!"=="Printer1EacLogoPath" set "P1_EAC_LOGO_PATH=!VAL!"
    if /I "!KEY!"=="Printer1EacHeightPx" set "P1_EAC_HEIGHT_PX=!VAL!"
    if /I "!KEY!"=="Printer1EacGapPx" set "P1_EAC_GAP_PX=!VAL!"
    if /I "!KEY!"=="Printer1EacOffsetX" set "P1_EAC_OFFSET_X=!VAL!"
    if /I "!KEY!"=="Printer1EacOffsetY" set "P1_EAC_OFFSET_Y=!VAL!"
    if /I "!KEY!"=="Printer2Label" set "P2_LABEL=!VAL!"
    if /I "!KEY!"=="Printer2X" set "P2_X=!VAL!"
    if /I "!KEY!"=="Printer2Y" set "P2_Y=!VAL!"
    if /I "!KEY!"=="Printer2ModuleSize" set "P2_MODULE_SIZE=!VAL!"
    if /I "!KEY!"=="Printer2TextOffsetY" set "P2_TEXT_OFFSET_Y=!VAL!"
    if /I "!KEY!"=="Printer2TextWidthPx" set "P2_TEXT_WIDTH_PX=!VAL!"
    if /I "!KEY!"=="Printer2GtinFontSize" set "P2_GTIN_FONT_SIZE=!VAL!"
    if /I "!KEY!"=="Printer2NameFontSize" set "P2_NAME_FONT_SIZE=!VAL!"
    if /I "!KEY!"=="Printer2Ai21FontSize" set "P2_AI21_FONT_SIZE=!VAL!"
    if /I "!KEY!"=="Printer2LineGap" set "P2_LINE_GAP=!VAL!"
    if /I "!KEY!"=="Printer2NameMaxLines" set "P2_NAME_MAX_LINES=!VAL!"
    if /I "!KEY!"=="Printer2FontPath" set "P2_FONT_PATH=!VAL!"
    if /I "!KEY!"=="Printer2EacLogoPath" set "P2_EAC_LOGO_PATH=!VAL!"
    if /I "!KEY!"=="Printer2EacHeightPx" set "P2_EAC_HEIGHT_PX=!VAL!"
    if /I "!KEY!"=="Printer2EacGapPx" set "P2_EAC_GAP_PX=!VAL!"
    if /I "!KEY!"=="Printer2EacOffsetX" set "P2_EAC_OFFSET_X=!VAL!"
    if /I "!KEY!"=="Printer2EacOffsetY" set "P2_EAC_OFFSET_Y=!VAL!"
  )
)
goto :eof

:save_ini
> "%INI%" echo ; LABEL PRINTING TOOL v3.6 CONFIGURATION / КОНФИГУРАЦИЯ ПЕЧАТИ ЭТИКЕТОК v3.6
>> "%INI%" echo ; Paths are relative to the folder containing the BAT file / Пути относительны папке, где лежит BAT-файл
>> "%INI%" echo ; Lines starting with ; are comments / Строки, начинающиеся с ;, являются комментариями
>> "%INI%" echo.
>> "%INI%" echo [Settings]
>> "%INI%" echo Language=%LANG%
>> "%INI%" echo ExcelFile=%EXCEL%
>> "%INI%" echo BarcodeColumn=%COLUMN%
>> "%INI%" echo GtinColumn=%GTIN_COLUMN%
>> "%INI%" echo NameColumn=%NAME_COLUMN%
>> "%INI%" echo Printer1IP=%PRINTER1%
>> "%INI%" echo Printer2IP=%PRINTER2%
>> "%INI%" echo CurrentIP=%IP%
>> "%INI%" echo LastRow=%LAST_ROW%
>> "%INI%" echo.
>> "%INI%" echo [Printer1_60x30]
>> "%INI%" echo Printer1Label=%P1_LABEL%
>> "%INI%" echo Printer1X=%P1_X%
>> "%INI%" echo Printer1Y=%P1_Y%
>> "%INI%" echo Printer1ModuleSize=%P1_MODULE_SIZE%
>> "%INI%" echo Printer1TextOffsetY=%P1_TEXT_OFFSET_Y%
>> "%INI%" echo Printer1TextWidthPx=%P1_TEXT_WIDTH_PX%
>> "%INI%" echo Printer1GtinFontSize=%P1_GTIN_FONT_SIZE%
>> "%INI%" echo Printer1NameFontSize=%P1_NAME_FONT_SIZE%
>> "%INI%" echo Printer1Ai21FontSize=%P1_AI21_FONT_SIZE%
>> "%INI%" echo Printer1LineGap=%P1_LINE_GAP%
>> "%INI%" echo Printer1NameMaxLines=%P1_NAME_MAX_LINES%
>> "%INI%" echo Printer1FontPath=%P1_FONT_PATH%
>> "%INI%" echo Printer1EacLogoPath=%P1_EAC_LOGO_PATH%
>> "%INI%" echo Printer1EacHeightPx=%P1_EAC_HEIGHT_PX%
>> "%INI%" echo Printer1EacGapPx=%P1_EAC_GAP_PX%
>> "%INI%" echo Printer1EacOffsetX=%P1_EAC_OFFSET_X%
>> "%INI%" echo Printer1EacOffsetY=%P1_EAC_OFFSET_Y%
>> "%INI%" echo.
>> "%INI%" echo [Printer2_60x40]
>> "%INI%" echo Printer2Label=%P2_LABEL%
>> "%INI%" echo Printer2X=%P2_X%
>> "%INI%" echo Printer2Y=%P2_Y%
>> "%INI%" echo Printer2ModuleSize=%P2_MODULE_SIZE%
>> "%INI%" echo Printer2TextOffsetY=%P2_TEXT_OFFSET_Y%
>> "%INI%" echo Printer2TextWidthPx=%P2_TEXT_WIDTH_PX%
>> "%INI%" echo Printer2GtinFontSize=%P2_GTIN_FONT_SIZE%
>> "%INI%" echo Printer2NameFontSize=%P2_NAME_FONT_SIZE%
>> "%INI%" echo Printer2Ai21FontSize=%P2_AI21_FONT_SIZE%
>> "%INI%" echo Printer2LineGap=%P2_LINE_GAP%
>> "%INI%" echo Printer2NameMaxLines=%P2_NAME_MAX_LINES%
>> "%INI%" echo Printer2FontPath=%P2_FONT_PATH%
>> "%INI%" echo Printer2EacLogoPath=%P2_EAC_LOGO_PATH%
>> "%INI%" echo Printer2EacHeightPx=%P2_EAC_HEIGHT_PX%
>> "%INI%" echo Printer2EacGapPx=%P2_EAC_GAP_PX%
>> "%INI%" echo Printer2EacOffsetX=%P2_EAC_OFFSET_X%
>> "%INI%" echo Printer2EacOffsetY=%P2_EAC_OFFSET_Y%
goto :eof
