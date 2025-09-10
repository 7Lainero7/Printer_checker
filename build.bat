@echo off
chcp 65001 >nul
echo 🚀 Запуск автоматической сборки QR Printer Check
echo.

REM Активируем виртуальное окружение
if exist ".venv\Scripts\activate" (
    call .venv\Scripts\activate
    echo ✅ Виртуальное окружение активировано
)

REM Запускаем сборку
python build.py

echo.
pause