# Импортируем необходимые библиотеки
import pandas as pd  # Для работы с данными и Excel файлами
import glob  # Для поиска всех файлов xlsx в папке
import os  # Для работы с путями к файлам
import sys  # Для работы с кодировкой

# Настройка кодировки для консоли Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

# ========================================
# ШАГ 1: НАСТРОЙКИ
# ========================================
# Указываем путь к папке с файлами
folder_path = r"r:\Подразделения\РИСК-менеджмент\Внутренние\1 - РЫНОЧНЫЙ РИСК\ТОРГОВА КНИГА\0-2026\01-02-2026\ovdp_data"

# Название выходного файла
output_file = "fair_prise_ovdp.xlsx"

# ========================================
# ШАГ 2: ПОИСК ВСЕХ XLSX ФАЙЛОВ В ПАПКЕ
# ========================================
# Создаем шаблон для поиска: все файлы с расширением .xlsx
search_pattern = os.path.join(folder_path, "*.xlsx")

# Получаем список всех xlsx файлов в папке
all_files = glob.glob(search_pattern)

print(f"Найдено файлов: {len(all_files)}")

# ========================================
# ШАГ 3: ЧТЕНИЕ И ПЕРЕИМЕНОВАНИЕ ПЕРВОГО СТОЛБЦА В КАЖДОМ ФАЙЛЕ
# ========================================
# Создаем пустой список для хранения данных из каждого файла
dataframes_list = []

# Проходим по каждому найденному файлу
for file in all_files:
    # Читаем Excel файл (первая строка - заголовки)
    df = pd.read_excel(file)
    
    # КЛЮЧЕВОЕ ИЗМЕНЕНИЕ: переименовываем первый столбец (индекс 0) в 'Date'
    # независимо от его текущего названия
    df.rename(columns={df.columns[0]: 'Date'}, inplace=True)
    
    # Добавляем прочитанные данные в список
    dataframes_list.append(df)
    print(f"Прочитан файл: {os.path.basename(file)}")

# Объединяем все данные в один DataFrame
# ignore_index=True - сбрасываем старые индексы и создаем новые
combined_df = pd.concat(dataframes_list, ignore_index=True)

print(f"\nВсего строк после объединения: {len(combined_df)}")

# ========================================
# ШАГ 4: ВЫБОР И ПЕРЕИМЕНОВАНИЕ СТОЛБЦОВ
# ========================================
# Создаем словарь соответствия: старое название -> новое название
columns_mapping = {
    'Date': 'Date',  # Первый столбец уже переименован в 'Date' на предыдущем шаге
    'ISIN': 'ISIN',
    'Валюта номіналу цінного папера': 'CUR',
    'Справедлива вартість одного цінного папера з урахуванням накопиченого купонного доходу, у валюті номіналу': 'Cost',
    'Дохідність до погашення, %': 'Yield',
    'Дата погашення': 'MaturityDate'
}

# Выбираем только нужные столбцы и переименовываем их
result_df = combined_df[list(columns_mapping.keys())].copy()
result_df.columns = list(columns_mapping.values())

# ========================================
# ШАГ 5: ПРЕОБРАЗОВАНИЕ ДАТ
# ========================================
# Преобразуем столбцы с датами в формат datetime
# Это необходимо для правильного расчета разницы между датами
result_df['Date'] = pd.to_datetime(result_df['Date'])
result_df['MaturityDate'] = pd.to_datetime(result_df['MaturityDate'])

# ========================================
# ШАГ 6: РАСЧЕТ СТОЛБЦА TERM
# ========================================
# Вычисляем разницу в днях между датой погашения и текущей датой
# .dt.days извлекает количество дней из объекта timedelta
result_df['days_diff'] = (result_df['MaturityDate'] - result_df['Date']).dt.days

# Функция для определения категории Term на основе количества дней
def calculate_term(days):
    """
    Определяет категорию срока (Term) на основе количества дней.
    
    Параметры:
    days (int): Количество дней между Date и MaturityDate
    
    Возвращает:
    str: Категория срока (1M, 3M, 6M, 9M, 1Y, 2Y, 5Y, 10Y, UP10Y)
    """
    if days <= 31:
        return '1M'
    elif days <= 92:
        return '3M'
    elif days <= 183:
        return '6M'
    elif days <= 365:
        return '1Y'
    elif days <= 729:
        return '2Y'
    elif days <= 1094:
        return '3Y'
    elif days <= 1459:
        return '4Y'
    elif days <= 1824:
        return '5Y'
    elif days <= 2554:
        return '7Y'
    elif days <= 3649:
        return '10Y'
    elif days <= 5474:
        return '15Y'
    elif days <= 7299:
        return '20Y'    
    else:
        return 'UP20Y'

# Применяем функцию к каждой строке столбца days_diff
# apply() применяет функцию к каждому элементу столбца
result_df['Term'] = result_df['days_diff'].apply(calculate_term)

# Удаляем временный столбец days_diff (он нам больше не нужен)
result_df = result_df.drop('days_diff', axis=1)

# ========================================
# ШАГ 7: СОРТИРОВКА ПО СТОЛБЦУ DATE
# ========================================
# Сортируем DataFrame по столбцу Date от старых дат к новым (ascending=True)
# inplace=True - изменяем сам DataFrame, не создаем новый
result_df.sort_values(by='Date', ascending=True, inplace=True)

print(f"\n✓ Данные отсортированы по столбцу Date (от старых к новым)")

# ========================================
# ШАГ 8: СОХРАНЕНИЕ РЕЗУЛЬТАТА
# ========================================
# Формируем полный путь для выходного файла
output_path = os.path.join(folder_path, output_file)

# Сохраняем результат в Excel файл
# index=False - не сохраняем индексы строк в Excel
# engine='openpyxl' - указываем движок для работы с xlsx (для Python 3.12)
result_df.to_excel(output_path, index=False, engine='openpyxl')

print(f"\n✓ Готово! Файл сохранен: {output_path}")
print(f"Итоговое количество строк: {len(result_df)}")
print(f"Столбцы в файле: {list(result_df.columns)}")
