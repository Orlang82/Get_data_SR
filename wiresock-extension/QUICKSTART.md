# Быстрый старт - Wiresock Auto Launcher

## Установка за 5 минут

### 1. Сборка Native Host (один раз)

```bash
cd native-host
build.bat
```

Это создаст файл `wiresock_host.exe`.

### 2. Загрузка расширения в Chrome

1. Откройте Chrome
2. Перейдите на `chrome://extensions/`
3. Включите "Режим разработчика" (Developer mode) в правом верхнем углу
4. Нажмите "Загрузить распакованное расширение"
5. Выберите папку `wiresock-extension`
6. **ВАЖНО:** Скопируйте ID расширения (длинная строка из букв, например: `abcdefghijklmnopqrstuvwxyz123456`)

### 3. Настройка Native Host

1. Откройте файл `native-host/com.wiresock.launcher.json`
2. Замените `YOUR_EXTENSION_ID_HERE` на скопированный ID:

```json
{
  "name": "com.wiresock.launcher",
  "description": "Wiresock Auto Launcher Native Host",
  "path": "C:\\Users\\ORLANG_NOUTE\\OneDrive\\Python\\Get_data_SR\\wiresock-extension\\native-host\\wiresock_host.exe",
  "type": "stdio",
  "allowed_origins": [
    "chrome-extension://ВАШ_ID_РАСШИРЕНИЯ/"
  ]
}
```

### 4. Регистрация Native Host

Запустите **от имени администратора**:

```bash
cd native-host
install.bat
```

### 5. Настройка расширения

1. Нажмите на иконку расширения в Chrome
2. Укажите путь к конфигурации wiresock (например: `C:\configs\my-vpn.conf`)
3. Нажмите "Сохранить путь"
4. Добавьте домены для отслеживания (например: `mail.ru`)

### 6. Готово!

Теперь при переходе на добавленные домены автоматически запустится:
```
wiresock-client.exe run -config <ваш_путь> -log-level none
```

## Проверка работы

1. Нажмите кнопку "Тестировать подключение" в popup расширения
2. Проверьте уведомление Chrome
3. Посмотрите лог: `C:\Users\<ваше_имя>\wiresock_host.log`

## Важные пути (при необходимости изменить)

### Путь к wiresock-client.exe

По умолчанию: `C:\Program Files\WireSock VPN Client\bin\wiresock-client.exe`

Изменить в файле `native-host/wiresock_host.py`, строка ~48:
```python
wiresock_exe = r"C:\Program Files\WireSock VPN Client\bin\wiresock-client.exe"
```

После изменения пересоберите: `build.bat`

### Путь к wiresock_host.exe

Изменить в `native-host/com.wiresock.launcher.json`:
```json
"path": "C:\\путь\\к\\wiresock_host.exe"
```

## Типичные проблемы

### "Ошибка: Native host has exited"

1. Проверьте ID расширения в `com.wiresock.launcher.json`
2. Убедитесь, что `wiresock_host.exe` существует
3. Запустите `install.bat` от администратора

### Wiresock не запускается

1. Проверьте путь к `wiresock-client.exe` в `wiresock_host.py`
2. Проверьте путь к конфигурации в настройках расширения
3. Посмотрите лог: `C:\Users\<имя>\wiresock_host.log`

## Удаление

```bash
cd native-host
uninstall.bat
```

Затем удалите расширение в `chrome://extensions/`

---

Подробная документация в [README.md](README.md)
