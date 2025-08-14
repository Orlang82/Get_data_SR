import sys
# Настройка кодировки stdout/stderr для корректного вывода в Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
import shutil
import time
from datetime import datetime
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from winotify import Notification, audio

# Базовая директория для копирования файлов
DEST_BASE = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС"
# Путь к иконке для уведомлений
ICON_PATH = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС\СКРИПТЫ\PyScripts\Get_data_SR\watcher\icon.ico"
MAX_DEPTH = 5  # Не используется, можно удалить

# Конфигурация директорий для наблюдения и условий отбора файлов
WATCH_CONFIGS = [
    {
        "watch_dir": r"q:\STAT\new_stat\STA_ARCH\ARX_F01\2025", 
        "conditions": [
            lambda name: name.startswith("01x") and name.endswith(".xlsx"), 
            lambda name: name.startswith("норм")
        ]
    },
    {
        "watch_dir": r"q:\STAT\new_stat\STA_ARCH\ARX_FC5\2025", 
        "conditions": [
            lambda name: name.startswith(("c5", "с5"))
        ]
    },
    {
        "watch_dir": r"q:\STAT\new_stat\STA_ARCH\ARX_6RX\2025", 
        "conditions": [
            lambda name: name.startswith("6rx")
        ]
    },
    {
        "watch_dir": r"q:\STAT\new_stat\STA_ARCH\ARX_6JX\2025",
        "conditions": [
            lambda name: name.startswith("6jx") or name.startswith("активи")
        ]
    },
    {
        "watch_dir": r"q:\STAT\new_stat\STA_ARCH\ARX_F6KX\2025", 
        "conditions": 
        [
            lambda name: name.startswith("6kx") or name.startswith("sr")
        ]
    },
    {
        "watch_dir": r"q:\STAT\new_stat\STA_ARCH\ARX_F42\2025",
        "conditions": [
            lambda name: name.startswith("42x")
        ]
    }
]

def show_notification(file_name):
    # Показывает уведомление Windows о новом файле.
    toast = Notification(app_id="Stat Watcher", title="📊 Новый файл STAT", msg=file_name, icon=ICON_PATH)
    toast.set_audio(audio.Default, loop=False)
    toast.show()
    time.sleep(1)  # Короткая пауза для корректного отображения уведомления

class MultiDirHandler(FileSystemEventHandler):
    # Обработчик событий файловой системы для нескольких директорий.
    def __init__(self, conditions):
        super().__init__()
        self.conditions = conditions  # Список функций-условий для отбора файлов
        self.last_processed = {}  # Словарь для отслеживания уже обработанных файлов (file_path -> (size, time))

    def process_file(self, file_path: Path):
        # Проверяет файл по условиям и копирует его, если условия выполняются.
        try:
            file_stat = file_path.stat().st_size  # Получаем размер файла
        except FileNotFoundError:
            return  # Файл мог быть удалён до обработки

        now = time.time()

        # Проверка: если файл уже обрабатывался недавно и не изменился — пропускаем
        if file_path in self.last_processed:
            last_size, last_time = self.last_processed[file_path]
            if last_size == file_stat and (now - last_time) < 2:
                return

        self.last_processed[file_path] = (file_stat, now)

        file_name = file_path.name.lower()
        # Проверяем, удовлетворяет ли имя файла хотя бы одному условию
        if any(cond(file_name) for cond in self.conditions):
            show_notification(file_path.name)
            today_str = datetime.now().strftime("%d-%m-%Y")
            dest_dir = Path(DEST_BASE) / today_str
            dest_dir.mkdir(parents=True, exist_ok=True)

            try:
                shutil.copy2(file_path, dest_dir / file_path.name)
                print(f"[INFO] Файл {file_path.name} скопирован в {dest_dir}")
            except Exception as e:
                print(f"[ERROR] Не удалось скопировать {file_path.name}: {e}")

    def on_created(self, event):
        # Обработка события создания файла.
        if not event.is_directory:
            self.process_file(Path(event.src_path))

    def on_modified(self, event):
        # Обработка события изменения файла.
        if not event.is_directory:
            self.process_file(Path(event.src_path))

if __name__ == "__main__":
    # Запуск наблюдателей для всех директорий из конфигурации
    observers = []
    for config in WATCH_CONFIGS:
        path = config["watch_dir"]
        conditions = config["conditions"]
        handler = MultiDirHandler(conditions)
        observer = Observer()
        observer.schedule(handler, path, recursive=True)
        observer.start()
        observers.append(observer)
        print(f"[INFO] Запущено наблюдение за {path}")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        # Корректное завершение работы при нажатии Ctrl+C
        for obs in observers:
            obs.stop()
        for obs in observers:
            obs.join()
