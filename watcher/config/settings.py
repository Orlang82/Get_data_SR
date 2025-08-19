"""
Основные настройки и пути для файлового мониторинга
"""

import sys
from pathlib import Path

# Определяем базовую директорию проекта
if getattr(sys, 'frozen', False):
    # Если запущен как exe (скомпилированный)
    BASE_DIR = Path(sys.executable).parent
else:
    # Если запущен как .py скрипт
    BASE_DIR = Path(__file__).parent.parent

# Директория для логов (создается рядом с exe файлом на локальной машине)
LOG_DIR = BASE_DIR / "logs"

# Базовая директория для копирования файлов (сетевой диск)
DEST_BASE = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС"

# Путь к иконке для уведомлений (сетевой диск)
ICON_PATH = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС\СКРИПТЫ\PyScripts\Get_data_SR\watcher\icon.ico"

# Настройки для обработки файлов
COPY_DELAY = 3.0  # Увеличена задержка перед копированием для сетевых файлов
STABILITY_CHECK_INTERVAL = 1.0  # Увеличен интервал проверки стабильности
MAX_COPY_ATTEMPTS = 5  # Увеличено количество попыток копирования

# Интервалы мониторинга (в секундах)
HEARTBEAT_INTERVAL = 300  # 5 минут
HEALTH_CHECK_INTERVAL = 60  # 1 минута

# Глобальная переменная для контроля работы приложения
RUNNING = True

# Настройки логирования
LOG_MAX_BYTES = 10 * 1024 * 1024  # 10 MB
LOG_BACKUP_COUNT = 5

# Настройки уведомлений
NOTIFICATION_APP_ID = "Stat Watcher"
NOTIFICATION_TITLE = "📊 Новый файл STAT"

# Настройки для проверки свободного места (в байтах)
MIN_FREE_SPACE = 1024 * 1024 * 1024  # 1GB

# Коды сетевых ошибок Windows для retry логики
NETWORK_ERROR_CODES = [53, 59, 64, 67]
MAX_NETWORK_ERRORS = 3
