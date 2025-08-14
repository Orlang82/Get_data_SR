import sys
import logging
from pathlib import Path
from datetime import datetime
import signal
import os
import threading
import time

# Настройка кодировки stdout/stderr для корректного вывода в Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

import shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from winotify import Notification, audio

# Для системного трея
import pystray
from PIL import Image, ImageDraw

# Определяем директорию для логов (рядом с exe файлом)
if getattr(sys, 'frozen', False):
    # Если запущен как exe
    BASE_DIR = Path(sys.executable).parent
else:
    # Если запущен как .py скрипт
    BASE_DIR = Path(__file__).parent

LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

# Настройка ротации логов
from logging.handlers import RotatingFileHandler

# Настройка логирования с ротацией
log_handler = RotatingFileHandler(
    LOG_DIR / 'file_watcher.log',
    maxBytes=10*1024*1024,  # 10 MB
    backupCount=5,
    encoding='utf-8'
)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        log_handler,
    ]
)
logger = logging.getLogger(__name__)

# Глобальные переменные для контроля работы
RUNNING = True
app_icon = None

# Базовая директория для копирования файлов
DEST_BASE = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС"
# Путь к иконке для уведомлений
ICON_PATH = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС\СКРИПТЫ\PyScripts\Get_data_SR\watcher\icon.ico"

# Настройки для обработки файлов
COPY_DELAY = 3.0  # Увеличена задержка перед копированием для сетевых файлов
STABILITY_CHECK_INTERVAL = 1.0  # Увеличен интервал проверки стабильности
MAX_COPY_ATTEMPTS = 5  # Увеличено количество попыток копирования

def normalize_filename_for_comparison(filename):
    """
    Нормализует имя файла для сравнения, заменяя похожие кириллические 
    и латинские символы на единый вариант.
    """
    # Словарь замен: кириллица -> латиница
    replacements = {
        'а': 'a', 'А': 'A',
        'е': 'e', 'Е': 'E', 
        'к': 'k', 'К': 'K',
        'м': 'm', 'М': 'M',
        'н': 'n', 'Н': 'N',
        'о': 'o', 'О': 'O',
        'р': 'r', 'Р': 'R',
        'с': 'c', 'С': 'C',
        'т': 't', 'Т': 'T',
        'у': 'u', 'У': 'U',
        'х': 'x', 'Х': 'X',
        'ј': 'j',  # Сербская ј
    }
    
    result = filename.lower()
    for cyr, lat in replacements.items():
        result = result.replace(cyr, lat.lower())
    
    return result

def create_flexible_condition(patterns):
    """
    Создает условие, которое работает с различными вариантами 
    кириллицы/латиницы в именах файлов.
    """
    def condition(filename):
        normalized = normalize_filename_for_comparison(filename)
        return any(normalized.startswith(pattern.lower()) for pattern in patterns)
    return condition

# Конфигурация директорий для наблюдения и условий отбора файлов
WATCH_CONFIGS = [
    {
        "watch_dir": r"q:\STAT\new_stat\STA_ARCH\ARX_F01\2025", 
        "conditions": [
            lambda name: normalize_filename_for_comparison(name).startswith("01x") and name.endswith(".xlsx"), 
            lambda name: normalize_filename_for_comparison(name).startswith(("норм", "norm"))
        ]
    },
    {
        "watch_dir": r"q:\STAT\new_stat\STA_ARCH\ARX_FC5\2025", 
        "conditions": [
            create_flexible_condition(["c5", "с5"])
        ]
    },
    {
        "watch_dir": r"q:\STAT\new_stat\STA_ARCH\ARX_6RX\2025", 
        "conditions": [
            create_flexible_condition(["6rx", "6рх"])
        ]
    },
    {
        "watch_dir": r"q:\STAT\new_stat\STA_ARCH\ARX_6JX\2025",
        "conditions": [
            create_flexible_condition(["6jx", "6јх"]),
            lambda name: normalize_filename_for_comparison(name).startswith(("активи", "aktivi"))
        ]
    },
    {
        "watch_dir": r"q:\STAT\new_stat\STA_ARCH\ARX_F6KX\2025", 
        "conditions": [
            create_flexible_condition(["6kx", "6кх"]),
            lambda name: normalize_filename_for_comparison(name).startswith("sr")
        ]
    },
    {
        "watch_dir": r"q:\STAT\new_stat\STA_ARCH\ARX_F42\2025",
        "conditions": [
            create_flexible_condition(["42x"])
        ]
    }
]
def create_tray_icon():
    """Создает простую иконку для системного трея."""
    try:
        # Пытаемся загрузить существующую иконку
        if Path(ICON_PATH).exists():
            image = Image.open(ICON_PATH)
            if image.size != (64, 64):
                image = image.resize((64, 64), Image.Resampling.LANCZOS)
            return image
    except Exception as e:
        logger.debug(f"Не удалось загрузить иконку из {ICON_PATH}: {e}")
    
    # Создаем простую иконку программно
    image = Image.new('RGB', (64, 64), color='blue')
    draw = ImageDraw.Draw(image)
    # Рисуем простой символ папки
    draw.rectangle([10, 20, 54, 50], fill='lightblue', outline='darkblue', width=2)
    draw.rectangle([15, 15, 35, 25], fill='lightblue', outline='darkblue', width=2)
    # Добавляем символ "глаза" для мониторинга
    draw.ellipse([20, 30, 30, 40], fill='white', outline='black')
    draw.ellipse([23, 33, 27, 37], fill='black')
    return image

def show_notification(file_name):
    """Показывает уведомление Windows о новом файле."""
    try:
        toast = Notification(
            app_id="Stat Watcher", 
            title="📊 Новый файл STAT", 
            msg=file_name, 
            icon=ICON_PATH
        )
        toast.set_audio(audio.Default, loop=False)
        toast.show()
        time.sleep(0.1)
        logger.info(f"Показано уведомление для файла: {file_name}")
    except Exception as e:
        logger.warning(f"Не удалось показать уведомление для {file_name}: {e}")

def wait_for_file_stability(file_path: Path, max_wait_time=10):
    """
    Ожидает стабилизации файла (перестанет изменяться размер).
    Возвращает True если файл стабилен, False если превышено время ожидания.
    """
    start_time = time.time()
    previous_size = None
    network_error_count = 0
    max_network_errors = 3
    
    while time.time() - start_time < max_wait_time:
        try:
            current_size = file_path.stat().st_size
            network_error_count = 0  # Сбрасываем счетчик при успехе
            
            if previous_size is not None and previous_size == current_size:
                logger.debug(f"Файл {file_path.name} стабилен, размер: {current_size}")
                return True
            previous_size = current_size
            time.sleep(STABILITY_CHECK_INTERVAL)
            
        except FileNotFoundError:
            logger.warning(f"Файл {file_path} исчез во время ожидания стабилизации")
            return False
            
        except OSError as e:
            # Обрабатываем сетевые ошибки (коды 59, 53, 64, и др.)
            if e.winerror in [53, 59, 64, 67]:  # Различные сетевые ошибки Windows
                network_error_count += 1
                logger.warning(f"Сетевая ошибка при проверке {file_path.name} (попытка {network_error_count}/{max_network_errors}): {e}")
                
                if network_error_count >= max_network_errors:
                    logger.error(f"Превышено количество сетевых ошибок для {file_path.name}")
                    return False
                
                # Увеличиваем задержку при сетевых ошибках
                time.sleep(min(2 * network_error_count, 5))
                continue
            else:
                logger.error(f"Ошибка при проверке стабильности файла {file_path}: {e}")
                return False
                
        except Exception as e:
            logger.error(f"Неожиданная ошибка при проверке стабильности файла {file_path}: {e}")
            return False
    
    logger.warning(f"Превышено время ожидания стабилизации для файла {file_path.name}")
    return True

def copy_file_with_retries(src_path: Path, dest_path: Path, max_attempts=MAX_COPY_ATTEMPTS):
    """Копирует файл с несколькими попытками в случае неудачи."""
    for attempt in range(1, max_attempts + 1):
        try:
            # Проверяем существование файла с обработкой сетевых ошибок
            file_exists = False
            for check_attempt in range(3):
                try:
                    file_exists = src_path.exists()
                    break
                except OSError as e:
                    if e.winerror in [53, 59, 64, 67] and check_attempt < 2:
                        logger.debug(f"Сетевая ошибка при проверке существования файла, повтор...")
                        time.sleep(1)
                        continue
                    raise
            
            if not file_exists:
                logger.error(f"Исходный файл не существует: {src_path}")
                return False
                
            # Проверяем доступность файла для чтения с retry логикой
            read_success = False
            for read_attempt in range(3):
                try:
                    with open(src_path, 'rb') as f:
                        f.read(1)  # Читаем первый байт для проверки доступности
                    read_success = True
                    break
                except OSError as e:
                    if e.winerror in [53, 59, 64, 67] and read_attempt < 2:
                        logger.debug(f"Сетевая ошибка при чтении файла, повтор через {read_attempt + 1} сек...")
                        time.sleep(read_attempt + 1)
                        continue
                    raise
            
            if not read_success:
                logger.warning(f"Не удалось прочитать файл {src_path.name} для проверки доступности")
                if attempt < max_attempts:
                    time.sleep(2 * attempt)
                    continue
                return False
            
            # Копируем файл
            shutil.copy2(src_path, dest_path)
            logger.info(f"Файл {src_path.name} успешно скопирован в {dest_path.parent}")
            return True
            
        except PermissionError as e:
            logger.warning(f"Попытка {attempt}/{max_attempts}: Нет доступа к файлу {src_path.name}: {e}")
            if attempt < max_attempts:
                time.sleep(1 * attempt)
        except FileNotFoundError as e:
            logger.error(f"Попытка {attempt}/{max_attempts}: Файл не найден {src_path.name}: {e}")
            if attempt < max_attempts:
                time.sleep(0.5 * attempt)
        except OSError as e:
            # Специальная обработка сетевых ошибок
            if e.winerror in [53, 59, 64, 67]:
                logger.warning(f"Попытка {attempt}/{max_attempts}: Сетевая ошибка при копировании {src_path.name}: {e}")
                if attempt < max_attempts:
                    time.sleep(3 * attempt)  # Больше времени для сетевых ошибок
            else:
                logger.error(f"Попытка {attempt}/{max_attempts}: Ошибка ОС при копировании {src_path.name}: {e}")
                if attempt < max_attempts:
                    time.sleep(1 * attempt)
        except Exception as e:
            logger.error(f"Попытка {attempt}/{max_attempts}: Неожиданная ошибка при копировании {src_path.name}: {e}")
            if attempt < max_attempts:
                time.sleep(1 * attempt)
    
    logger.error(f"Не удалось скопировать файл {src_path.name} после {max_attempts} попыток")
    return False

class MultiDirHandler(FileSystemEventHandler):
    """Обработчик событий файловой системы для нескольких директорий."""
    
    def __init__(self, conditions):
        super().__init__()
        self.conditions = conditions
        self.processed_files = set()
        self.pending_files = {}

    def should_process_file(self, file_path: Path):
        """Проверяет, нужно ли обрабатывать файл."""
        file_name_lower = file_path.name.lower()
        return any(cond(file_name_lower) for cond in self.conditions)

    def process_file(self, file_path: Path, event_type="unknown"):
        """Обрабатывает файл: проверяет условия и копирует."""
        file_key = str(file_path)
        
        if file_key in self.processed_files:
            return
            
        if not self.should_process_file(file_path):
            return

        logger.info(f"Обнаружен файл для обработки: {file_path.name} (событие: {event_type})")
        
        # Ждем стабилизации файла
        time.sleep(COPY_DELAY)
        
        if not wait_for_file_stability(file_path):
            return
            
        # Показываем уведомление
        show_notification(file_path.name)
        
        # Создаем директорию назначения
        today_str = datetime.now().strftime("%d-%m-%Y")
        dest_dir = Path(DEST_BASE) / today_str
        try:
            dest_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            logger.error(f"Не удалось создать директорию {dest_dir}: {e}")
            return

        # Копируем файл
        dest_path = dest_dir / file_path.name
        if copy_file_with_retries(file_path, dest_path):
            self.processed_files.add(file_key)
            
            if len(self.processed_files) > 1000:
                self.processed_files.clear()
                logger.info("Очищен список обработанных файлов")

    def on_created(self, event):
        """Обработка события создания файла."""
        if not event.is_directory:
            self.process_file(Path(event.src_path), "created")

    def on_modified(self, event):
        """Обработка события изменения файла."""
        if not event.is_directory:
            file_path = Path(event.src_path)
            if self.should_process_file(file_path):
                def delayed_process():
                    time.sleep(2)
                    self.process_file(file_path, "modified")
                
                threading.Thread(target=delayed_process, daemon=True).start()

def test_filename_conditions():
    """Тестирует условия на примерах ваших файлов."""
    test_files = [
        "6КХ_13082025.xlsx",
        "6КХ_дані_13082025.xlsx", 
        "sr_13082025.TXT",
        "6kx_test.xlsx",
        "С5_13082025.xlsx",
        "c5_test.xlsx",
        "01X_13082025.xlsx",
        "нормативы_13082025.xlsx",
        "6RX_13082025.xlsx",
        "6рх_test.xlsx",
        "активи вкл до файлу_13082025.xlsx",
        "6JX_13082025.xlsx",
        "42x_test.xlsx"
    ]
    
    logger.info("=== Тестирование условий фильтрации файлов ===")
    
    for config in WATCH_CONFIGS:
        watch_dir = config["watch_dir"]
        conditions = config["conditions"]
        logger.info(f"\nДиректория: {watch_dir}")
        
        matching_files = []
        for test_file in test_files:
            if any(cond(test_file.lower()) for cond in conditions):
                matching_files.append(test_file)
        
        if matching_files:
            logger.info(f"  Подходящие файлы: {', '.join(matching_files)}")
        else:
            logger.info("  Подходящих файлов не найдено")
    
    logger.info("=== Конец тестирования ===\n")
def validate_paths():
    """Проверяет существование всех путей из конфигурации."""
    logger.info("Проверка существования директорий...")
    
    dest_path = Path(DEST_BASE)
    if not dest_path.exists():
        logger.warning(f"Целевая директория не существует: {dest_path}")
    
    for i, config in enumerate(WATCH_CONFIGS):
        watch_path = Path(config["watch_dir"])
        if not watch_path.exists():
            logger.warning(f"Директория для наблюдения не существует: {watch_path}")
        else:
            logger.info(f"✓ Директория найдена: {watch_path}")

def monitor_observer_health(observers):
    """Мониторит состояние наблюдателей и перезапускает их при необходимости."""
    for i, (observer, config) in enumerate(observers):
        try:
            if not observer.is_alive():
                logger.warning(f"Наблюдатель {i+1} не активен. Попытка перезапуска...")
                
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except:
                    pass
                
                path = config["watch_dir"]
                conditions = config["conditions"]
                handler = MultiDirHandler(conditions)
                new_observer = Observer()
                new_observer.schedule(handler, path, recursive=True)
                new_observer.start()
                
                observers[i] = (new_observer, config)
                logger.info(f"Наблюдатель {i+1} перезапущен для: {path}")
                
        except Exception as e:
            logger.error(f"Ошибка при мониторинге наблюдателя {i+1}: {e}")

def daemon_heartbeat():
    """Периодическая проверка работоспособности."""
    try:
        free_space = shutil.disk_usage(DEST_BASE).free
        if free_space < 1024 * 1024 * 1024:  # Меньше 1GB
            logger.warning(f"Мало свободного места на диске: {free_space / (1024**3):.2f} GB")
        
        logger.debug("Процесс работает нормально")
        return True
    except Exception as e:
        logger.warning(f"Проблема при проверке состояния: {e}")
        return False

# Системный трей
class TrayApp:
    def __init__(self):
        self.observers = []
        self.icon = None
        
    def setup_tray(self):
        """Настройка иконки в системном трее."""
        image = create_tray_icon()
        
        menu = pystray.Menu(
            pystray.MenuItem("Статус", self.show_status),
            pystray.MenuItem("Открыть логи", self.open_logs),
            pystray.MenuItem("Перезапустить", self.restart_watchers),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Выход", self.quit_app)
        )
        
        self.icon = pystray.Icon("FileWatcher", image, "File Watcher - Мониторинг файлов", menu)
        
    def show_status(self, icon, item):
        """Показывает статус работы."""
        active_watchers = sum(1 for obs, _ in self.observers if obs.is_alive())
        total_watchers = len(self.observers)
        
        toast = Notification(
            app_id="Stat Watcher",
            title="📊 File Watcher - Статус",
            msg=f"Активных наблюдателей: {active_watchers}/{total_watchers}",
            icon=ICON_PATH
        )
        toast.show()
        
    def open_logs(self, icon, item):
        """Открывает папку с логами."""
        try:
            os.startfile(LOG_DIR)
        except Exception as e:
            logger.error(f"Не удалось открыть папку с логами: {e}")
            
    def restart_watchers(self, icon, item):
        """Перезапускает всех наблюдателей."""
        logger.info("Инициирован перезапуск наблюдателей...")
        monitor_observer_health(self.observers)
        
        toast = Notification(
            app_id="Stat Watcher",
            title="📊 File Watcher",
            msg="Наблюдатели перезапущены",
            icon=ICON_PATH
        )
        toast.show()
        
    def quit_app(self, icon, item):
        """Завершает работу приложения."""
        global RUNNING
        logger.info("Инициировано завершение работы из трея...")
        RUNNING = False
        if icon:
            icon.stop()

def main():
    """Основная функция приложения."""
    global RUNNING, app_icon
    
    logger.info("=== Запуск фонового процесса мониторинга файлов ===")
    
    # Создаем приложение трея
    tray_app = TrayApp()
    tray_app.setup_tray()
    
    # Проверяем пути
    validate_paths()
    
    # Запускаем наблюдателей
    observers = []
    
    for config in WATCH_CONFIGS:
        path = config["watch_dir"]
        conditions = config["conditions"]
        
        if not Path(path).exists():
            logger.warning(f"Пропуск несуществующей директории: {path}")
            continue
            
        handler = MultiDirHandler(conditions)
        observer = Observer()
        observer.schedule(handler, path, recursive=True)
        observer.start()
        observers.append((observer, config))
        logger.info(f"✓ Запущено наблюдение за: {path}")

    if not observers:
        logger.error("Не удалось запустить ни одного наблюдателя")
        return 1
    else:
        tray_app.observers = observers
        
        logger.info(f"Процесс запущен в фоне. Активных наблюдателей: {len(observers)}")
        logger.info("Иконка добавлена в системный трей")
        
        # Запускаем основной цикл в отдельном потоке
        def background_loop():
            global RUNNING  # Добавляем объявление глобальной переменной
            heartbeat_counter = 0
            HEARTBEAT_INTERVAL = 300  # 5 минут
            HEALTH_CHECK_INTERVAL = 60  # 1 минута
            
            # Основной цикл службы
            try:
                while RUNNING:
                    time.sleep(1)
                    heartbeat_counter += 1
                    
                    if heartbeat_counter % HEALTH_CHECK_INTERVAL == 0:
                        monitor_observer_health(observers)
                    
                    if heartbeat_counter % HEARTBEAT_INTERVAL == 0:
                        daemon_heartbeat()
                        heartbeat_counter = 0
            
            except KeyboardInterrupt:
                logger.info("Получен сигнал KeyboardInterrupt")
            finally:
                RUNNING = False
                logger.info("Остановка наблюдателей...")
                for observer, config in observers:
                    try:
                        observer.stop()
                        observer.join(timeout=10)
                        logger.debug(f"Наблюдатель остановлен: {config['watch_dir']}")
                    except Exception as e:
                        logger.warning(f"Ошибка при остановке: {e}")
            
            logger.info("Фоновый процесс мониторинга файлов завершен")
        
        # Запускаем фоновый поток
        bg_thread = threading.Thread(target=background_loop, daemon=True)
        bg_thread.start()
        
        try:
            # Показываем уведомление о запуске
            toast = Notification(
                app_id="Stat Watcher",
                title="📊 File Watcher",
                msg="Фоновый мониторинг файлов запущен",
                icon=ICON_PATH
            )
            toast.show()
            
            # Запускаем системный трей (блокирующий вызов)
            if tray_app.icon:
                tray_app.icon.run()
            
        except KeyboardInterrupt:
            logger.info("Получен сигнал KeyboardInterrupt")
        finally:
            RUNNING = False
            logger.info("Фоновый процесс мониторинга файлов завершен")

if __name__ == "__main__":
    main()