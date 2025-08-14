import os
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path
from collections import defaultdict
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from winotify import Notification, audio

# === НАСТРОЙКИ ===
WATCH_DIR = r"q:\STAT\new_stat\STA_ARCH\ARX_F01\2025"
DEST_BASE = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС"
MAX_DEPTH = 5
ICON_PATH = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС\СКРИПТЫ\PyScripts\Get_data_SR\watcher\icon.ico"

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
    time.sleep(0.5)

# === Обработчик событий ===
class StatFileHandler(FileSystemEventHandler):
    last_processed = defaultdict(float)  # хранит время последней обработки файла

    def process_file(self, file_path: Path):
        now = time.time()
        if now - self.last_processed[file_path] < 2:
            return  # защита от дублей в пределах 2 секунд
        self.last_processed[file_path] = now

        file_name = file_path.name.lower()
        if (file_name.startswith("01x") and file_name.endswith(".xlsx")) or file_name.startswith("норм"):
            show_notification(file_path.name)

            today_str = datetime.now().strftime("%d-%m-%Y")
            dest_dir = Path(DEST_BASE) / today_str
            dest_dir.mkdir(parents=True, exist_ok=True)

            try:
                shutil.copy2(file_path, dest_dir / file_path.name)
                print(f"[INFO] Файл {file_path.name} скопирован в {dest_dir}")
            except Exception as e:
                print(f"[ERROR] Не удалось скопировать {file_path.name}: {e}")

    def on_any_event(self, event):
        if not event.is_directory and event.event_type in ("created", "modified"):
            self.process_file(Path(event.src_path))

# === Запуск слежения ===
def start_watching():
    event_handler = StatFileHandler()
    observer = Observer()

    for root, dirs, _ in os.walk(WATCH_DIR):
        depth = Path(root).relative_to(WATCH_DIR).parts
        if len(depth) <= MAX_DEPTH:
            observer.schedule(event_handler, root, recursive=False)

    observer.start()
    print(f"[INFO] Слежение запущено за {WATCH_DIR} (глубина {MAX_DEPTH})")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        import ctypes
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    start_watching()