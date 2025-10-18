@echo off
chcp 65001 >nul
REM Сборка исполняемого файла из Python скрипта с помощью PyInstaller

setlocal

echo ============================================
echo Сборка Wiresock Host в исполняемый файл
echo ============================================
echo.

REM Проверка наличия PyInstaller
python -m pip show pyinstaller >nul 2>&1

if %ERRORLEVEL% NEQ 0 (
    echo PyInstaller не установлен. Установка...
    python -m pip install pyinstaller

    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo ОШИБКА: Не удалось установить PyInstaller
        pause
        exit /b 1
    )
)

echo.
echo Сборка исполняемого файла...
echo.

REM Сборка с PyInstaller
python -m PyInstaller ^
    --onefile ^
    --noconsole ^
    --name wiresock_host ^
    --clean ^
    wiresock_host.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ============================================
    echo Сборка завершена успешно!
    echo ============================================
    echo.
    echo Исполняемый файл: dist\wiresock_host.exe
    echo.
    echo Скопируйте dist\wiresock_host.exe в текущую папку
    echo или обновите путь в com.wiresock.launcher.json
    echo.

    REM Автоматическое копирование
    if exist "dist\wiresock_host.exe" (
        echo Копирование wiresock_host.exe в текущую папку...
        copy /Y "dist\wiresock_host.exe" "wiresock_host.exe"
        echo Готово!
    )
) else (
    echo.
    echo ОШИБКА: Сборка не удалась
    echo.
)

pause
