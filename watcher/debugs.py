import sys
# Настройка кодировки stdout/stderr для корректного вывода в Windows
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

DEST_BASE = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС"
ICON_PATH = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС\СКРИПТЫ\PyScripts\Get_data_SR\watcher\icon.ico"

WATCH_CONFIGS = [
    {
        "watch_dir": r"q:\STAT\new_stat\STA_ARCH\ARX_FC5\2025",
        "conditions": [lambda name: name.startswith(("c5", "с5"))]
    }
]

def show_notification(file_name):
    toast = Notification(
        app_id="Stat Watcher",
        title="📊 Новый файл STAT",
        msg=file_name,
        icon=ICON_PATH
    )
    toast.set_audio(audio.Default, loop=False)
    toast.show()
    time.sleep(0.5)

class DebugHandler(FileSystemEventHandler):
    def __init__(self, conditions):
        super().__init__()
        self.conditions = conditions
        self.last_processed = {}  # file_path -> (size, time)

    def process_file(self, file_path: Path):
        try:
            file_stat = file_path.stat().st_size
        except FileNotFoundError:
            return

        now = time.time()
        if file_path in self.last_processed:
            last_size, last_time = self.last_processed[file_path]
            if last_size == file_stat and (now - last_time) < 2:
                print(f"[DEBUG] Пропущено (дубликат): {file_path.name}")
                return

        self.last_processed[file_path] = (file_stat, now)

        file_name = file_path.name.lower()

        # Логируем отладочную информацию
        first_two = file_name[:2]
        ord_values = [ord(ch) for ch in first_two]
        print(f"[DEBUG] Обрабатываем файл: {file_path.name}")
        print(f"[DEBUG] Первые 2 символа: '{first_two}' -> ord: {ord_values}")

        for cond in self.conditions:
            try:
                result = cond(file_name)
                print(f"[DEBUG] Условие {cond}: {result}")
            except Exception as e:
                print(f"[ERROR] Ошибка в условии {cond}: {e}")

        if any(cond(file_name) for cond in self.conditions):
            print(f"[INFO] Файл {file_path.name} прошёл проверку")
            show_notification(file_path.name)
            today_str = datetime.now().strftime("%d-%m-%Y")
            dest_dir = Path(DEST_BASE) / today_str
            dest_dir.mkdir(parents=True, exist_ok=True)

            try:
                shutil.copy2(file_path, dest_dir / file_path.name)
                print(f"[INFO] Файл {file_path.name} скопирован в {dest_dir}")
            except Exception as e:
                print(f"[ERROR] Не удалось скопировать {file_path.name}: {e}")
        else:
            print(f"[DEBUG] Файл {file_path.name} НЕ прошёл условия")

    def on_created(self, event):
        if not event.is_directory:
            self.process_file(Path(event.src_path))

    def on_modified(self, event):
        if not event.is_directory:
            self.process_file(Path(event.src_path))

if __name__ == "__main__":
    observers = []

    for config in WATCH_CONFIGS:
        path = config["watch_dir"]
        conditions = config["conditions"]
        handler = DebugHandler(conditions)
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
