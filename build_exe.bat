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
echo 2. Проверка синтаксиса Python...
python -m py_compile navigator_app.py
if %errorlevel% neq 0 (
    echo [ОШИБКА] В коде navigator_app.py обнаружены синтаксические ошибки!
    pause
    exit /b %errorlevel%
)
echo [OK] Синтаксис проверен успешно.

if not "%~1"=="" (
    echo [ИНФО] Указана версия для сборки: %~1
    python -c "import re; v = re.sub(r'^[vV]', '', '%~1'); c = open('navigator_app.py', encoding='utf-8').read(); open('navigator_app.py', 'w', encoding='utf-8').write(re.sub(r'APP_VERSION\s*=\s*[\"\\\'][^\"\\\']+[\"\\\']', f'APP_VERSION = \"{v}\"', c)); print(f'Установлена версия APP_VERSION = {v}')"
)

echo.
echo 3. Компиляция navigator_app.py в автономный EXE (PyInstaller)...
if exist "app.ico" (
    echo [ИНФО] Обнаружена иконка app.ico, собираем с пользовательской иконкой...
    pyinstaller --noconsole --onefile --icon "app.ico" --add-data "app.ico;." --name "NavigatorApp" navigator_app.py
) else (
    pyinstaller --noconsole --onefile --name "NavigatorApp" navigator_app.py
)
if %errorlevel% neq 0 (
    echo [ОШИБКА] Сборка завершилась с ошибкой.
    pause
    exit /b %errorlevel%
)

echo.
echo 4. Копирование файлов шаблонов (Excel, config и LICENSE) к EXE...
if not exist "dist" mkdir "dist"
if not exist "dist\logs" mkdir "dist\logs"
if exist "event_list.xlsx" copy "event_list.xlsx" "dist\event_list.xlsx" > nul
if exist "programm_list.xlsx" copy "programm_list.xlsx" "dist\programm_list.xlsx" > nul
if exist "study_list.xlsx" copy "study_list.xlsx" "dist\study_list.xlsx" > nul
if exist "list.xlsx" copy "list.xlsx" "dist\list.xlsx" > nul
if exist "config.example.json" copy "config.example.json" "dist\config.example.json" > nul
if exist "config.json" copy "config.json" "dist\config.json" > nul
if exist "LICENSE" copy "LICENSE" "dist\LICENSE" > nul
if exist "app.ico" copy "app.ico" "dist\app.ico" > nul

echo.
echo ========================================================
echo Готово!
echo Исполняемый файл и шаблоны находятся в папке: dist/
echo   - dist\NavigatorApp.exe
echo   - dist\event_list.xlsx
echo   - dist\programm_list.xlsx
echo   - dist\study_list.xlsx
echo   - dist\config.example.json
echo   - dist\config.json (если был создан)
echo   - dist\LICENSE
echo   - dist\logs\
echo ========================================================
echo.
pause
