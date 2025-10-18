#!/usr/bin/env python3
"""
Скрипт для создания простых иконок для расширения Wiresock Auto Launcher
"""

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("Ошибка: необходима библиотека Pillow")
    print("Установите: pip install pillow")
    exit(1)

def create_icon(size, filename):
    """Создает иконку заданного размера"""

    # Создаем изображение с синим фоном
    img = Image.new('RGB', (size, size), color='#2196F3')
    draw = ImageDraw.Draw(img)

    # Рисуем зеленый круг в центре
    margin = size // 4
    draw.ellipse(
        [margin, margin, size - margin, size - margin],
        fill='#4CAF50',
        outline='white',
        width=max(1, size // 32)
    )

    # Добавляем текст "W" (для Wiresock)
    try:
        # Пытаемся загрузить системный шрифт
        font_size = size // 2
        try:
            font = ImageFont.truetype("arial.ttf", font_size)
        except:
            try:
                font = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", font_size)
            except:
                font = ImageFont.load_default()
    except:
        font = ImageFont.load_default()

    text = "W"

    # Получаем размеры текста
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]

    # Центрируем текст
    position = (
        (size - text_width) // 2,
        (size - text_height) // 2 - size // 8
    )

    draw.text(position, text, fill='white', font=font)

    # Сохраняем изображение
    img.save(filename, 'PNG')
    print(f"[OK] Создана иконка: {filename} ({size}x{size})")

def main():
    """Создает все необходимые иконки"""

    print("Создание иконок для Wiresock Auto Launcher...")
    print()

    sizes = [
        (16, 'icon16.png'),
        (48, 'icon48.png'),
        (128, 'icon128.png')
    ]

    for size, filename in sizes:
        try:
            create_icon(size, filename)
        except Exception as e:
            print(f"[ERROR] Ошибка создания {filename}: {e}")

    print()
    print("Готово! Иконки созданы.")
    print("Теперь можно загрузить расширение в браузер.")

if __name__ == '__main__':
    main()
