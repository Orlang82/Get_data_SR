import sys
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
import shutil
import time
from datetime import datetime
from pathlib import Path
from collections import defaultdict
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from winotify import Notification, audio

# === НАСТРОЙКИ ===
DEST_BASE = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС"
ICON_PATH = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС\СКРИПТЫ\PyScripts\Get_data_SR\watcher\icon.ico"
MAX_DEPTH = 5

# Список конфигураций: путь + условия
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
        "conditions": [
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

# === Уведомление Windows 10/11 ===
def show_notification(file_name):
    toast = Notification(
        app_id="Stat Watcher",
        title="📊 Новый файл STAT",
        msg=file_name,
        icon=ICON_PATH
    )
    toast.set_audio(audio.Default, loop=False)
    toast.show()
    time.sleep(1)

# === Обработчик событий ===
class MultiDirHandler(FileSystemEventHandler):
    def __init__(self, conditions):
        super().__init__()
        self.conditions = conditions
        self.last_processed = defaultdict(float)

    def process_file(self, file_path: Path):
        now = time.time()
        if now - self.last_processed[file_path] < 2:
            return
        self.last_processed[file_path] = now

        file_name = file_path.name.lower()
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
        if not event.is_directory:
            self.process_file(Path(event.src_path))

    def on_modified(self, event):
        if not event.is_directory:
            self.process_file(Path(event.src_path))

# === Запуск наблюдателей ===
if __name__ == "__main__":
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
        for obs in observers:
            obs.stop()
        for obs in observers:
            obs.join()
