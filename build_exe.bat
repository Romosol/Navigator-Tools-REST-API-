@echo off
chcp 65001 > nul
echo ========================================================
echo Сборка NavigatorApp.exe для Windows (PyInstaller)
echo ========================================================
echo.

echo 1. Установка необходимых библиотек...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ОШИБКА] Не удалось установить зависимости. Убедитесь, что Python и pip установлены.
    pause
    exit /b %errorlevel%
)

echo.
echo 2. Компиляция navigator_app.py в автономный EXE...
pyinstaller --noconsole --onefile --name "NavigatorApp" navigator_app.py
if %errorlevel% neq 0 (
    echo [ОШИБКА] Сборка завершилась с ошибкой.
    pause
    exit /b %errorlevel%
)

echo.
echo 3. Копирование файлов шаблонов (Excel и config.json) к EXE...
if not exist "dist" mkdir "dist"
if not exist "dist\logs" mkdir "dist\logs"
if exist "event_list.xlsx" copy "event_list.xlsx" "dist\event_list.xlsx" > nul
if exist "programm_list.xlsx" copy "programm_list.xlsx" "dist\programm_list.xlsx" > nul
if exist "study_list.xlsx" copy "study_list.xlsx" "dist\study_list.xlsx" > nul
if exist "list.xlsx" copy "list.xlsx" "dist\list.xlsx" > nul
if exist "config.json" copy "config.json" "dist\config.json" > nul

echo.
echo ========================================================
echo Готово!
echo Исполняемый файл и шаблоны находятся в папке: dist/
echo   - dist\NavigatorApp.exe
echo   - dist\event_list.xlsx
echo   - dist\programm_list.xlsx
echo   - dist\study_list.xlsx
echo   - dist\config.json
echo   - dist\logs\
echo ========================================================
echo.
pause
