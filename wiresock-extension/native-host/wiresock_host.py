#!/usr/bin/env python3
"""
Native Messaging Host для Wiresock Auto Launcher
Принимает сообщения от Chrome Extension и запускает wiresock-client.exe
"""

import sys
import json
import struct
import subprocess
import logging
import os
from pathlib import Path

# Настройка логирования
log_path = Path.home() / 'wiresock_host.log'
logging.basicConfig(
    filename=str(log_path),
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def send_message(message):
    """
    Отправляет сообщение обратно в расширение Chrome
    Native Messaging использует специальный формат:
    - 4 байта: длина сообщения (uint32)
    - N байт: JSON сообщение
    """
    try:
        encoded_message = json.dumps(message).encode('utf-8')
        message_length = len(encoded_message)

        # Отправляем длину сообщения (4 байта, little-endian)
        sys.stdout.buffer.write(struct.pack('I', message_length))
        # Отправляем само сообщение
        sys.stdout.buffer.write(encoded_message)
        sys.stdout.buffer.flush()

        logging.info(f"Отправлено сообщение: {message}")
    except Exception as e:
        logging.error(f"Ошибка отправки сообщения: {e}")

def read_message():
    """
    Читает сообщение от расширения Chrome
    """
    try:
        # Читаем длину сообщения (4 байта)
        raw_length = sys.stdin.buffer.read(4)

        if len(raw_length) == 0:
            logging.info("Поток ввода закрыт")
            return None

        message_length = struct.unpack('I', raw_length)[0]

        # Читаем само сообщение
        message = sys.stdin.buffer.read(message_length).decode('utf-8')

        logging.info(f"Получено сообщение: {message}")
        return json.loads(message)
    except Exception as e:
        logging.error(f"Ошибка чтения сообщения: {e}")
        return None

def launch_wiresock(config_path, domain):
    """
    Запускает wiresock-client.exe с указанными параметрами
    """
    try:
        # Путь к wiresock-client.exe
        # Можно настроить или определить автоматически
        wiresock_exe = r"c:\Program Files\WireSock Secure Connect\bin\wiresock-client.exe"

        # Проверяем существование исполняемого файла
        if not os.path.exists(wiresock_exe):
            # Попытка найти в PATH
            wiresock_exe = "wiresock-client.exe"

        # Формируем команду
        command = [
            wiresock_exe,
            'run',
            '-config', config_path,
            '-log-level', 'none'
        ]

        logging.info(f"Запуск команды: {' '.join(command)}")

        # Запускаем процесс в фоновом режиме
        # Используем CREATE_NO_WINDOW для Windows, чтобы не показывать окно консоли
        startupinfo = None
        if sys.platform == 'win32':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            startupinfo=startupinfo,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
        )

        logging.info(f"Wiresock запущен для домена {domain}, PID: {process.pid}")

        return {
            'success': True,
            'message': f'Wiresock успешно запущен для {domain}',
            'pid': process.pid
        }

    except FileNotFoundError:
        error_msg = f"wiresock-client.exe не найден. Проверьте путь: {wiresock_exe}"
        logging.error(error_msg)
        return {
            'success': False,
            'error': error_msg
        }
    except Exception as e:
        error_msg = f"Ошибка запуска wiresock: {str(e)}"
        logging.error(error_msg)
        return {
            'success': False,
            'error': error_msg
        }

def main():
    """
    Основная функция - читает сообщения и обрабатывает их
    """
    logging.info("Native Messaging Host запущен")

    try:
        while True:
            message = read_message()

            if message is None:
                break

            # Обрабатываем сообщение
            action = message.get('action')

            if action == 'launch':
                domain = message.get('domain', 'unknown')
                config_path = message.get('configPath', '')

                if not config_path:
                    response = {
                        'success': False,
                        'error': 'Не указан путь к конфигурации'
                    }
                elif not os.path.exists(config_path):
                    response = {
                        'success': False,
                        'error': f'Файл конфигурации не найден: {config_path}'
                    }
                else:
                    response = launch_wiresock(config_path, domain)

                send_message(response)
            else:
                send_message({
                    'success': False,
                    'error': f'Неизвестное действие: {action}'
                })

    except Exception as e:
        logging.error(f"Критическая ошибка: {e}")
        send_message({
            'success': False,
            'error': str(e)
        })

    logging.info("Native Messaging Host завершен")

if __name__ == '__main__':
    main()
