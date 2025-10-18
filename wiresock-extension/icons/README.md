# Иконки для расширения

Для корректной работы расширения необходимо создать три иконки:

- **icon16.png** - 16x16 пикселей (отображается в адресной строке)
- **icon48.png** - 48x48 пикселей (используется в chrome://extensions/)
- **icon128.png** - 128x128 пикселей (для Chrome Web Store и в деталях расширения)

## Быстрое создание иконок

### Вариант 1: Онлайн генераторы
1. Используйте сервис [Favicon Generator](https://realfavicongenerator.net/)
2. Загрузите любое изображение (например, логотип VPN)
3. Скачайте набор иконок разных размеров
4. Переименуйте файлы в icon16.png, icon48.png, icon128.png

### Вариант 2: Простые цветные квадраты (для тестирования)
Создайте простые PNG файлы нужных размеров с помощью любого графического редактора:
- Photoshop
- GIMP (бесплатно)
- Paint.NET (бесплатно)
- Онлайн редактор [Pixlr](https://pixlr.com/)

Рекомендуемые цвета для VPN-тематики:
- Синий: #2196F3
- Зеленый: #4CAF50
- Фиолетовый: #9C27B0

### Вариант 3: Использование Python PIL

Создайте файл `create_icons.py` в этой папке:

```python
from PIL import Image, ImageDraw, ImageFont

def create_icon(size, filename):
    # Создаем изображение с градиентом
    img = Image.new('RGB', (size, size), color='#2196F3')
    draw = ImageDraw.Draw(img)

    # Рисуем круг
    margin = size // 4
    draw.ellipse([margin, margin, size-margin, size-margin],
                 fill='#4CAF50', outline='white', width=2)

    # Добавляем текст "W" (для Wiresock)
    try:
        font_size = size // 2
        font = ImageFont.truetype("arial.ttf", font_size)
    except:
        font = ImageFont.load_default()

    text = "W"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]

    position = ((size - text_width) // 2, (size - text_height) // 2 - size // 8)
    draw.text(position, text, fill='white', font=font)

    img.save(filename)
    print(f"Создана иконка: {filename}")

# Создаем все три иконки
create_icon(16, 'icon16.png')
create_icon(48, 'icon48.png')
create_icon(128, 'icon128.png')
```

Затем запустите:
```bash
pip install pillow
python create_icons.py
```

## Важно

После создания иконок убедитесь, что файлы:
1. Имеют правильные размеры (16x16, 48x48, 128x128 пикселей)
2. Сохранены в формате PNG
3. Находятся в папке `wiresock-extension/icons/`

## Рекомендации по дизайну

- Используйте простые, узнаваемые символы
- Иконка должна хорошо выглядеть на светлом и темном фоне
- Для VPN можно использовать символы: замок, щит, земной шар с замком
- Избегайте мелких деталей - они плохо видны на размере 16x16
