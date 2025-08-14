import os
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path
from collections import defaultdict
from threading import Thread, Event
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from winotify import Notification, audio

# === НАСТРОЙКИ ===
WATCH_DIR = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС\test"
DEST_BASE = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС\out"
MAX_DEPTH = 5
ICON_PATH = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС\СКРИПТЫ\PyScripts\Get_data_SR\watcher\icon.ico"

# === Буфер уведомлений ===
notification_buffer = set()
stop_event = Event()

def show_notification(file_list):
    """Показывает одно уведомление со списком файлов"""
    files_text = "\n".join(file_list)
    toast = Notification(
        app_id="Stat Watcher",
        title="📊 Новые файлы STAT",
        msg=files_text,
        icon=ICON_PATH
    )
    toast.set_audio(audio.Default, loop=False)
    toast.show()
    time.sleep(0.5)

def notification_worker():
    """Фоновая задача, которая раз в 5 секунд показывает уведомления"""
    while not stop_event.is_set():
        if notification_buffer:
            files_to_show = list(notification_buffer)
            notification_buffer.clear()
            show_notification(files_to_show)
        time.sleep(5)

# === Обработчик событий ===
class StatFileHandler(FileSystemEventHandler):
    last_processed = defaultdict(float)  # защита от дублей

    def process_file(self, file_path: Path):
        now = time.time()
        if now - self.last_processed[file_path] < 2:
            return
        self.last_processed[file_path] = now

        file_name = file_path.name.lower()
        if (file_name.startswith("01x") and file_name.endswith(".xlsx")) or file_name.startswith("норм"):
            # добавляем в буфер для группового уведомления
            notification_buffer.add(file_path.name)

            # копируем сразу
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

    # Запускаем поток для уведомлений
    notif_thread = Thread(target=notification_worker, daemon=True)
    notif_thread.start()

    observer.start()
    print(f"[INFO] Слежение запущено за {WATCH_DIR} (глубина {MAX_DEPTH})")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        stop_event.set()
        observer.stop()
    observer.join()

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        import ctypes
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    start_watching()
