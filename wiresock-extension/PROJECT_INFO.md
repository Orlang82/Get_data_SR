# Wiresock Auto Launcher - Информация о проекте

## Описание

**Wiresock Auto Launcher** - расширение для браузеров на основе Chromium, которое автоматически запускает `wiresock-client.exe` при обращении к определенным доменам через адресную строку или переход по ссылкам.

## Архитектура решения

### Компоненты

1. **Chrome Extension** (расширение браузера)
   - Отслеживает навигацию пользователя
   - Проверяет URL на соответствие списку доменов
   - Отправляет команды запуска через Native Messaging API
   - Предоставляет UI для управления настройками

2. **Native Messaging Host** (нативное приложение)
   - Python-скрипт, скомпилированный в .exe
   - Принимает сообщения от расширения
   - Запускает wiresock-client.exe с указанными параметрами
   - Ведет логирование операций

### Схема работы

```
Пользователь переходит на mail.ru
         ↓
Chrome Extension (background.js)
  ├─ webNavigation.onBeforeNavigate
  ├─ Проверка домена в списке
  └─ Если найден →
         ↓
Chrome Native Messaging API
  └─ chrome.runtime.sendNativeMessage()
         ↓
Native Host (wiresock_host.exe)
  ├─ Получение сообщения
  ├─ Валидация параметров
  └─ Запуск wiresock-client.exe
         ↓
wiresock-client.exe run -config <path> -log-level none
```

## Структура файлов

```
wiresock-extension/
│
├── manifest.json                    # Манифест Chrome Extension (Manifest V3)
├── background.js                    # Service Worker - основная логика
├── popup.html                       # UI управления доменами
├── popup.js                         # Логика UI
├── popup.css                        # Стили UI
│
├── icons/                           # Иконки расширения
│   ├── icon16.png                  # 16x16px
│   ├── icon48.png                  # 48x48px
│   ├── icon128.png                 # 128x128px
│   ├── create_icons.py             # Генератор иконок
│   └── README.md                   # Инструкция по иконкам
│
├── native-host/                     # Native Messaging Host
│   ├── wiresock_host.py            # Python скрипт хоста
│   ├── com.wiresock.launcher.json  # Манифест Native Host
│   ├── build.bat                   # Сборка .exe через PyInstaller
│   ├── install.bat                 # Регистрация в реестре Windows
│   └── uninstall.bat               # Удаление из реестра
│
├── README.md                        # Полная документация
├── QUICKSTART.md                    # Быстрый старт
└── PROJECT_INFO.md                  # Этот файл
```

## Технологии

### Chrome Extension
- **Manifest Version**: 3 (последняя версия)
- **APIs используемые**:
  - `chrome.webNavigation` - отслеживание переходов
  - `chrome.storage` - хранение настроек
  - `chrome.runtime` - Native Messaging
  - `chrome.notifications` - уведомления
- **Permissions**:
  - `storage` - сохранение списка доменов
  - `webNavigation` - перехват навигации
  - `nativeMessaging` - связь с хост-приложением
  - `host_permissions: <all_urls>` - доступ ко всем URL

### Native Host
- **Язык**: Python 3
- **Компиляция**: PyInstaller (в standalone .exe)
- **Протокол**: Native Messaging (stdin/stdout, JSON + length-prefix)
- **Логирование**: Файл в домашней директории пользователя

### Безопасность
- Whitelist расширений через `allowed_origins` в манифесте хоста
- Валидация путей и параметров перед запуском команд
- Запуск wiresock в фоновом режиме (скрытое окно)
- Проверка существования конфигурационных файлов

## Особенности реализации

### 1. Отслеживание навигации
Используется `chrome.webNavigation.onBeforeNavigate` вместо `webRequest`, так как:
- Не требует разрешения `webRequestBlocking` (удалено в MV3)
- Срабатывает раньше фактической загрузки страницы
- Работает только для главных фреймов (frameId === 0)

### 2. Native Messaging
Протокол общения:
```javascript
// Расширение → Хост
{
  "action": "launch",
  "domain": "mail.ru",
  "configPath": "C:\\path\\to\\config.conf"
}

// Хост → Расширение
{
  "success": true,
  "message": "Wiresock успешно запущен",
  "pid": 12345
}
```

### 3. Хранение данных
```javascript
chrome.storage.sync.set({
  domains: ["mail.ru", "example.com"],
  configPath: "C:\\configs\\vpn.conf"
})
```

Используется `storage.sync` для синхронизации между устройствами.

### 4. Service Worker (background.js)
В Manifest V3 фоновые скрипты заменены на Service Workers:
- Не выполняются постоянно, а активируются по событиям
- Сохраняют состояние в `chrome.storage`
- Более эффективное использование ресурсов

## Кастомизация

### Изменить путь к wiresock-client.exe

Файл: `native-host/wiresock_host.py`, строка ~48:

```python
wiresock_exe = r"C:\ВАШ\ПУТЬ\wiresock-client.exe"
```

После изменения запустить `build.bat`.

### Добавить дополнительные параметры запуска

Файл: `native-host/wiresock_host.py`, функция `launch_wiresock()`:

```python
command = [
    wiresock_exe,
    'run',
    '-config', config_path,
    '-log-level', 'none',
    # Добавьте свои параметры
    '-your-param', 'value'
]
```

### Изменить логику проверки доменов

Файл: `background.js`, функция `isDomainMonitored()`:

```javascript
function isDomainMonitored(url) {
    // Ваша логика проверки
}
```

## Совместимость

### Браузеры
- Google Chrome 88+
- Microsoft Edge 88+
- Opera 74+
- Brave
- Любой браузер на основе Chromium с поддержкой Manifest V3

### Операционные системы
- Windows 10/11 (основная поддержка)
- Windows 7/8.1 (возможно, требуется тестирование)

### Python
- Python 3.7+ (для разработки и сборки)
- После сборки Python не требуется (standalone .exe)

## Разработка

### Требования для разработки
```bash
pip install pillow pyinstaller
```

### Отладка расширения
1. Открыть `chrome://extensions/`
2. Найти "Wiresock Auto Launcher"
3. Нажать "Inspect views: background page"
4. Использовать консоль Chrome DevTools

### Отладка Native Host
Проверить лог файл:
```
C:\Users\<имя_пользователя>\wiresock_host.log
```

### Внесение изменений
1. Отредактировать файлы
2. Если изменен `wiresock_host.py`:
   ```bash
   cd native-host
   build.bat
   ```
3. Перезагрузить расширение в браузере (кнопка "Reload")

## Возможные улучшения

1. **UI**
   - Добавить индикатор активности в иконку расширения
   - История запусков
   - Настройки уведомлений

2. **Функциональность**
   - Автоматическое отключение при закрытии вкладки
   - Поддержка нескольких конфигураций для разных доменов
   - Экспорт/импорт настроек

3. **Безопасность**
   - Шифрование путей к конфигурациям
   - Подтверждение перед запуском (опционально)

4. **Мониторинг**
   - Просмотр логов прямо из расширения
   - Статистика использования
   - Проверка состояния процесса wiresock

## Лицензия

Открытый код, свободное использование.

## Автор

Создано с помощью Claude (Anthropic)

## Дата создания

Октябрь 2025
