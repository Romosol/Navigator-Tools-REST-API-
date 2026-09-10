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
echo 3. Копирование файлов шаблонов (list.xlsx и config.json) к EXE...
if not exist "dist" mkdir "dist"
if exist "list.xlsx" copy "list.xlsx" "dist\list.xlsx" > nul
if exist "config.json" copy "config.json" "dist\config.json" > nul

echo.
echo ========================================================
echo Готово!
echo Исполняемый файл и шаблоны находятся в папке: dist/
echo   - dist\NavigatorApp.exe
echo   - dist\list.xlsx
echo   - dist\config.json
echo ========================================================
echo.
pause
