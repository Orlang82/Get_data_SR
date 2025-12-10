import requests
from datetime import datetime, timedelta
import os
import sys
from pathlib import Path

# Настройка кодировки для консоли Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')
    # Альтернативный вариант, если верхний не сработает
    # locale.setpreferredencoding('UTF-8')

# ===== НАЛАШТУВАННЯ =====
# Вкажіть діапазон дат для скачування
START_DATE = datetime(2025, 11, 1)  # Початкова дата
END_DATE = datetime(2025, 12, 1)  # Кінцева дата
OUTPUT_FOLDER = r"r:\Подразделения\РИСК-менеджмент\Внутренние\1 - РЫНОЧНЫЙ РИСК\ТОРГОВА КНИГА\0-2025\01-12-2025\ovdp_data"  # Папка для збереження файлів

# ===== КОД СКРИПТУ =====
def download_ovdp_files(start_date, end_date, output_folder):
    """
    Скачує файли справедливої вартості ОВДП з сайту НБУ
    
    Параметри:
    - start_date: початкова дата (datetime)
    - end_date: кінцева дата (datetime)
    - output_folder: папка для збереження файлів
    """
    
    # Створюємо папку, якщо її немає
    Path(output_folder).mkdir(parents=True, exist_ok=True)
    
    # Лічильники для статистики
    downloaded = 0
    skipped = 0
    errors = 0
    
    # Перебираємо всі дні в діапазоні
    current_date = start_date
    while current_date <= end_date:
        # Формуємо URL за шаблоном
        year_month = current_date.strftime("%Y%m")  # Наприклад: 202501
        full_date = current_date.strftime("%Y%m%d")  # Наприклад: 20250130
        url = f"https://bank.gov.ua/files/Fair_value/{year_month}/{full_date}_fv.xlsx"
        
        # Формуємо ім'я файлу для збереження
        filename = os.path.join(output_folder, f"{full_date}_fv.xlsx")
        
        # Пропускаємо, якщо файл вже існує
        if os.path.exists(filename):
            print(f"⏭️  Пропущено (вже існує): {full_date}")
            skipped += 1
            current_date += timedelta(days=1)
            continue
        
        # Намагаємося скачати файл
        try:
            response = requests.get(url, timeout=10)
            
            # Перевіряємо, чи файл існує (код 200 = успішно)
            if response.status_code == 200:
                # Зберігаємо файл
                with open(filename, 'wb') as f:
                    f.write(response.content)
                print(f"✅ Скачано: {full_date}")
                downloaded += 1
            else:
                # Файл не знайдено (ймовірно, вихідний день)
                print(f"⚠️  Не знайдено: {full_date} (код: {response.status_code})")
                errors += 1
                
        except Exception as e:
            print(f"❌ Помилка при скачуванні {full_date}: {e}")
            errors += 1
        
        # Переходимо до наступного дня
        current_date += timedelta(days=1)
    
    # Виводимо підсумкову статистику
    print("\n" + "="*50)
    print(f"📊 СТАТИСТИКА:")
    print(f"   ✅ Скачано нових файлів: {downloaded}")
    print(f"   ⏭️  Пропущено (вже існують): {skipped}")
    print(f"   ⚠️  Не знайдено/помилки: {errors}")
    print(f"   📁 Файли збережено в папці: {output_folder}")
    print("="*50)

# ===== ЗАПУСК СКРИПТУ =====
if __name__ == "__main__":
    print("🚀 Початок скачування файлів ОВДП...")
    print(f"📅 Період: з {START_DATE.strftime('%d.%m.%Y')} по {END_DATE.strftime('%d.%m.%Y')}")
    print(f"📁 Папка збереження: {OUTPUT_FOLDER}\n")
    
    download_ovdp_files(START_DATE, END_DATE, OUTPUT_FOLDER)
    
    print("\n✅ Скрипт завершив роботу!")
