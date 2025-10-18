# Wiresock Auto Launcher

Расширение для браузеров на основе Chromium (Chrome, Edge, Opera, Brave и др.), которое автоматически запускает `wiresock-client.exe` при посещении указанных доменов.

## Возможности

- Автоматический запуск wiresock-client при переходе на отслеживаемые домены
- Управление списком доменов через удобный интерфейс
- Настройка пути к файлу конфигурации
- Уведомления о запуске подключений
- Поддержка всех браузеров на основе Chromium

## Структура проекта

```
wiresock-extension/
├── manifest.json           # Манифест расширения
├── background.js          # Фоновый скрипт (основная логика)
├── popup.html            # Интерфейс управления
├── popup.js              # Логика интерфейса
├── popup.css             # Стили интерфейса
├── icons/                # Иконки расширения (нужно добавить)
│   ├── icon16.png
│   ├── icon48.png
│   └── icon128.png
└── native-host/          # Native Messaging Host
    ├── wiresock_host.py          # Python скрипт хоста
    ├── com.wiresock.launcher.json # Манифест хоста
    ├── build.bat                  # Сборка исполняемого файла
    ├── install.bat                # Установка хоста в систему
    └── uninstall.bat              # Удаление хоста
```

## Установка

### Шаг 1: Подготовка иконок

Создайте иконки для расширения или используйте готовые:
- `icons/icon16.png` - 16x16 пикселей
- `icons/icon48.png` - 48x48 пикселей
- `icons/icon128.png` - 128x128 пикселей

Можно использовать любые PNG изображения нужного размера.

### Шаг 2: Сборка Native Host

1. Установите Python 3.x (если не установлен)
2. Перейдите в папку `native-host`
3. Запустите `build.bat` для создания исполняемого файла
4. После сборки файл `wiresock_host.exe` будет создан

```batch
cd native-host
build.bat
```

### Шаг 3: Установка расширения в браузер

1. Откройте Chrome/Edge/другой Chromium браузер
2. Перейдите на страницу расширений:
   - Chrome: `chrome://extensions/`
   - Edge: `edge://extensions/`
   - Opera: `opera://extensions/`
3. Включите "Режим разработчика" (Developer mode)
4. Нажмите "Загрузить распакованное расширение"
5. Выберите папку `wiresock-extension`
6. **ВАЖНО:** Скопируйте ID установленного расширения

### Шаг 4: Настройка Native Host

1. Откройте файл `native-host/com.wiresock.launcher.json`
2. Замените `YOUR_EXTENSION_ID_HERE` на скопированный ID расширения:

```json
{
  "name": "com.wiresock.launcher",
  "description": "Wiresock Auto Launcher Native Host",
  "path": "C:\\Users\\ORLANG_NOUTE\\OneDrive\\Python\\Get_data_SR\\wiresock-extension\\native-host\\wiresock_host.exe",
  "type": "stdio",
  "allowed_origins": [
    "chrome-extension://abcdefghijklmnopqrstuvwxyz123456/"
  ]
}
```

3. Проверьте путь к `wiresock_host.exe` - он должен быть абсолютным
4. При необходимости отредактируйте путь к `wiresock-client.exe` в файле `wiresock_host.py` (по умолчанию: `C:\Program Files\WireSock VPN Client\bin\wiresock-client.exe`)

### Шаг 5: Регистрация Native Host

Запустите `install.bat` **от имени администратора** для регистрации Native Host в системе:

```batch
cd native-host
install.bat
```

## Использование

### Первоначальная настройка

1. Нажмите на иконку расширения в браузере
2. В поле "Путь к конфигурации" укажите полный путь к вашему файлу конфигурации wiresock (например, `C:\configs\my-config.conf`)
3. Нажмите "Сохранить путь"

### Добавление доменов

1. Откройте popup расширения
2. В поле "Отслеживаемые домены" введите домен (например, `mail.ru`)
3. Нажмите "Добавить домен"

По умолчанию в списке уже есть примеры: `mail.ru` и `example.com`

### Автоматическая работа

После настройки расширение будет:
1. Отслеживать переходы по URL в браузере
2. Проверять, входит ли домен в список отслеживаемых
3. Автоматически запускать `wiresock-client run -config <путь> -log-level none`
4. Показывать уведомление об успешном запуске

## Настройка wiresock-client.exe

По умолчанию скрипт ищет `wiresock-client.exe` по пути:
```
C:\Program Files\WireSock VPN Client\bin\wiresock-client.exe
```

Если ваш путь отличается, отредактируйте переменную `wiresock_exe` в файле `native-host/wiresock_host.py` и пересоберите хост:

```python
wiresock_exe = r"C:\ВАШ\ПУТЬ\К\wiresock-client.exe"
```

## Логирование

Native Host ведет лог работы в файле:
```
C:\Users\<ваше_имя>\wiresock_host.log
```

В случае проблем проверьте этот файл для диагностики.

## Удаление

### Удаление расширения
1. Откройте страницу расширений в браузере
2. Найдите "Wiresock Auto Launcher"
3. Нажмите "Удалить"

### Удаление Native Host
Запустите `uninstall.bat` в папке `native-host`:

```batch
cd native-host
uninstall.bat
```

## Устранение неполадок

### Расширение не работает
1. Проверьте, что расширение включено в списке расширений
2. Откройте консоль расширения (chrome://extensions/ → Details → Inspect views: background page)
3. Проверьте наличие ошибок в консоли

### Native Host не отвечает
1. Проверьте, что `wiresock_host.exe` существует по указанному пути
2. Проверьте ID расширения в `com.wiresock.launcher.json`
3. Проверьте логи в `wiresock_host.log`
4. Убедитесь, что Native Host зарегистрирован в реестре:
   ```
   HKCU\SOFTWARE\Google\Chrome\NativeMessagingHosts\com.wiresock.launcher
   ```

### wiresock-client не запускается
1. Проверьте путь к `wiresock-client.exe` в `wiresock_host.py`
2. Проверьте путь к файлу конфигурации в настройках расширения
3. Убедитесь, что файл конфигурации существует и имеет правильный формат
4. Проверьте логи в `wiresock_host.log`

## Разработка

Для внесения изменений:
1. Отредактируйте нужные файлы
2. Если изменили `wiresock_host.py`, пересоберите: `build.bat`
3. Перезагрузите расширение в браузере (кнопка "Обновить" на странице расширений)

## Безопасность

- Расширение использует Native Messaging API для взаимодействия с системой
- Native Host запускается только по запросу от расширения
- Проверяется ID расширения для предотвращения несанкционированного доступа
- Все настройки хранятся в Chrome Storage (синхронизируются между устройствами)

## Лицензия

Свободное использование

## Поддержка

При возникновении проблем проверьте:
1. Логи расширения (background page console)
2. Логи Native Host (`wiresock_host.log`)
3. Правильность всех путей в конфигурации
