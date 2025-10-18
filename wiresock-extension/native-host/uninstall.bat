@echo off
chcp 65001 >nul
REM Удаление Native Messaging Host для Wiresock Auto Launcher

setlocal

set HOST_NAME=com.wiresock.launcher

echo ============================================
echo Удаление Wiresock Auto Launcher Native Host
echo ============================================
echo.

REM Удаление из HKLM (для всех пользователей)
echo Удаление из реестра для всех пользователей...
echo.

REM Google Chrome
echo [HKLM] Удаление для Google Chrome...
reg delete "HKLM\SOFTWARE\Google\Chrome\NativeMessagingHosts\%HOST_NAME%" /f 2>nul

REM Chromium
echo [HKLM] Удаление для Chromium...
reg delete "HKLM\SOFTWARE\Chromium\NativeMessagingHosts\%HOST_NAME%" /f 2>nul

REM Comet
echo [HKLM] Удаление для Comet...
reg delete "HKLM\SOFTWARE\Comet\NativeMessagingHosts\%HOST_NAME%" /f 2>nul

REM Opera
echo [HKLM] Удаление для Opera...
reg delete "HKLM\SOFTWARE\Opera Software\NativeMessagingHosts\%HOST_NAME%" /f 2>nul

echo.
echo Удаление из реестра для текущего пользователя...
echo.

REM Google Chrome - текущий пользователь
echo [HKCU] Удаление для Google Chrome...
reg delete "HKCU\SOFTWARE\Google\Chrome\NativeMessagingHosts\%HOST_NAME%" /f 2>nul

REM Chromium - текущий пользователь
echo [HKCU] Удаление для Chromium...
reg delete "HKCU\SOFTWARE\Chromium\NativeMessagingHosts\%HOST_NAME%" /f 2>nul

REM Comet - текущий пользователь
echo [HKCU] Удаление для Comet...
reg delete "HKCU\SOFTWARE\Comet\NativeMessagingHosts\%HOST_NAME%" /f 2>nul

REM Opera - текущий пользователь
echo [HKCU] Удаление для Opera...
reg delete "HKCU\SOFTWARE\Opera Software\NativeMessagingHosts\%HOST_NAME%" /f 2>nul

echo.
echo ============================================
echo Удаление завершено
echo ============================================
echo.

pause
