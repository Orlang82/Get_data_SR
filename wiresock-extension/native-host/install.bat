@echo off
chcp 65001 >nul
REM Установка Native Messaging Host для Wiresock Auto Launcher

setlocal

REM Определяем текущую директорию
set SCRIPT_DIR=%~dp0

REM Путь к манифесту
set MANIFEST_PATH=%SCRIPT_DIR%com.wiresock.launcher.json

REM Имя host
set HOST_NAME=com.wiresock.launcher

echo ============================================
echo Установка Wiresock Auto Launcher Native Host
echo ============================================
echo.

REM Проверка существования манифеста
if not exist "%MANIFEST_PATH%" (
    echo ОШИБКА: Манифест не найден: %MANIFEST_PATH%
    pause
    exit /b 1
)

echo Манифест найден: %MANIFEST_PATH%
echo.

REM Регистрация для всех пользователей (требует прав администратора)
echo Регистрация для ВСЕХ пользователей (требуется запуск от администратора)...
echo.

REM Google Chrome
echo [HKLM] Регистрация для Google Chrome...
reg add "HKLM\SOFTWARE\Google\Chrome\NativeMessagingHosts\%HOST_NAME%" /ve /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

REM Chromium
echo [HKLM] Регистрация для Chromium...
reg add "HKLM\SOFTWARE\Chromium\NativeMessagingHosts\%HOST_NAME%" /ve /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

REM Comet
echo [HKLM] Регистрация для Comet...
reg add "HKLM\SOFTWARE\Comet\NativeMessagingHosts\%HOST_NAME%" /ve /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

REM Opera
echo [HKLM] Регистрация для Opera...
reg add "HKLM\SOFTWARE\Opera Software\NativeMessagingHosts\%HOST_NAME%" /ve /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

echo.
echo Регистрация для текущего пользователя...
echo.

REM Google Chrome - текущий пользователь
echo [HKCU] Регистрация для Google Chrome...
reg add "HKCU\SOFTWARE\Google\Chrome\NativeMessagingHosts\%HOST_NAME%" /ve /t REG_SZ /d "%MANIFEST_PATH%" /f

REM Chromium - текущий пользователь
echo [HKCU] Регистрация для Chromium...
reg add "HKCU\SOFTWARE\Chromium\NativeMessagingHosts\%HOST_NAME%" /ve /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

REM Comet - текущий пользователь
echo [HKCU] Регистрация для Comet...
reg add "HKCU\SOFTWARE\Comet\NativeMessagingHosts\%HOST_NAME%" /ve /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

REM Opera - текущий пользователь
echo [HKCU] Регистрация для Opera...
reg add "HKCU\SOFTWARE\Opera Software\NativeMessagingHosts\%HOST_NAME%" /ve /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ============================================
    echo Установка завершена успешно!
    echo ============================================
    echo.
    echo Следующие шаги:
    echo 1. Убедитесь, что в манифесте указан правильный ID расширения
    echo 2. Создайте исполняемый файл из wiresock_host.py (см. build.bat)
    echo 3. Проверьте путь к wiresock-client.exe в wiresock_host.py
    echo.
) else (
    echo.
    echo ОШИБКА: Не удалось зарегистрировать native host
    echo.
)

pause
