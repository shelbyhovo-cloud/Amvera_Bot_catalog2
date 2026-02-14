"""
🕷️ GUI Приложение для парсинга товаров
Оконное приложение с кнопками и статус-баром
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import sys
import subprocess
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import io
import json
import zipfile

# Фикс кодировки для Windows консоли
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# ═══════════════════════════════════════════════════════════
# 📦 АВТОУСТАНОВКА ЗАВИСИМОСТЕЙ
# ═══════════════════════════════════════════════════════════

def install_dependencies():
    """Автоматически устанавливает необходимые пакеты."""
    required_packages = {
        'openpyxl': 'openpyxl==3.1.2',
        'requests': 'requests==2.31.0',
        'yfinance': 'yfinance',
    }

    missing_packages = []

    for package_name, package_spec in required_packages.items():
        try:
            __import__(package_name)
        except ImportError:
            missing_packages.append(package_spec)

    if missing_packages:
        print("📦 Устанавливаю недостающие зависимости...")
        print(f"   Пакеты: {', '.join(missing_packages)}")

        try:
            subprocess.check_call([
                sys.executable,
                '-m',
                'pip',
                'install',
                *missing_packages
            ])
            print("✅ Зависимости установлены успешно!\n")
        except subprocess.CalledProcessError as e:
            print(f"❌ Ошибка установки зависимостей: {e}")
            sys.exit(1)

# Устанавливаем зависимости
install_dependencies()

import openpyxl
import requests
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

# ═══════════════════════════════════════════════════════════
# 📄 СОЗДАНИЕ КРАСИВОГО ШАБЛОНА EXCEL
# ═══════════════════════════════════════════════════════════

def create_beautiful_template(file_path=None, brands=None):
    """Создаёт красиво оформленный шаблон Excel."""

    if file_path is None:
        script_dir = Path(__file__).parent
        file_path = script_dir / "products_links.xlsx"
    else:
        file_path = Path(file_path)

    wb = Workbook()
    ws = wb.active
    ws.title = "🛍 Товары"

    # Заголовки (без эмодзи, БЕЗ описания)
    headers = ["URL товара", "Название", "Цена (€)", "Группа", "Подгруппа", "Категория товара", "URL фото", "Локальное фото", "Размеры", "Последнее обновление", "Статус"]
    ws.append(headers)

    # Красивые стили для заголовков
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=14, name="Calibri")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Границы
    thin_border = Border(
        left=Side(style='thin', color='FFFFFF'),
        right=Side(style='thin', color='FFFFFF'),
        top=Side(style='thin', color='FFFFFF'),
        bottom=Side(style='thin', color='FFFFFF')
    )

    # Применяем стили к заголовкам
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border

    # Высота строки заголовка
    ws.row_dimensions[1].height = 30

    # Ширина колонок
    ws.column_dimensions['A'].width = 55  # URL товара
    ws.column_dimensions['B'].width = 35  # Название
    ws.column_dimensions['C'].width = 12  # Цена
    ws.column_dimensions['D'].width = 18  # Группа
    ws.column_dimensions['E'].width = 18  # Подгруппа
    ws.column_dimensions['F'].width = 20  # Категория товара
    ws.column_dimensions['G'].width = 45  # URL фото
    ws.column_dimensions['H'].width = 25  # Локальное фото
    ws.column_dimensions['I'].width = 25  # Размеры
    ws.column_dimensions['J'].width = 22  # Обновление
    ws.column_dimensions['K'].width = 18  # Статус

    # Примеры убраны - пустой шаблон
    examples = []

    # Цвета для строк (чередование)
    row_colors = ["F2F2F2", "FFFFFF"]

    # Стили для ячеек данных
    data_font = Font(size=11, name="Calibri")
    data_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    center_alignment = Alignment(horizontal="center", vertical="center")

    data_border = Border(
        left=Side(style='thin', color='D0D0D0'),
        right=Side(style='thin', color='D0D0D0'),
        top=Side(style='thin', color='D0D0D0'),
        bottom=Side(style='thin', color='D0D0D0')
    )

    for idx, row in enumerate(examples):
        row_num = idx + 2
        ws.append(row)

        # Цвет фона строки (по умолчанию)
        row_fill = PatternFill(start_color=row_colors[idx % 2], end_color=row_colors[idx % 2], fill_type="solid")

        # Заливки для конкретных столбцов
        name_fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")  # Серый для названия
        price_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")  # Зелёный для цены

        # Применяем стили к каждой ячейке
        for col_num, cell in enumerate(ws[row_num], start=1):
            # Специальные заливки для названия и цены
            if col_num == 2:  # B: Название
                cell.fill = name_fill
            elif col_num == 3:  # C: Цена
                cell.fill = price_fill
            else:
                cell.fill = row_fill

            cell.font = data_font
            cell.border = data_border

            # Выравнивание по центру для определённых колонок
            if col_num in [3, 9, 10, 11]:  # Цена, Размеры, Обновление, Статус
                cell.alignment = center_alignment
            else:
                cell.alignment = data_alignment

        # Высота строки
        ws.row_dimensions[row_num].height = 25

    # Закрепляем первую строку (заголовки)
    ws.freeze_panes = "A2"

    # Автофильтр (теперь до колонки S - включая расчетные столбцы)
    ws.auto_filter.ref = f"A1:S1"

    # ═══════════════════════════════════════════════════════════
    # 📊 ДОБАВЛЯЕМ РАСЧЕТНЫЕ СТОЛБЦЫ (L-S)
    # ═══════════════════════════════════════════════════════════

    calc_headers = [
        "Доставка (₽)",      # L
        "Закупка (₽)",       # M
        "Кэф Пети (%)",      # N
        "Наш Кэф (%)",       # O
        "Цена с дост. (₽)", # P
        "Цена без дост. (₽)", # Q
        "Наша Маржа (₽)",    # R
        "Маржа Пети (₽)"     # S
    ]

    # Оранжевое оформление для расчетных заголовков
    orange_header_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    calc_header_font = Font(bold=True, color="FFFFFF", size=12, name="Calibri")
    calc_header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    calc_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    for col_idx, header in enumerate(calc_headers, start=12):  # L=12
        cell = ws.cell(1, col_idx)
        cell.value = header
        cell.fill = orange_header_fill
        cell.font = calc_header_font
        cell.alignment = calc_header_alignment
        cell.border = calc_border

        # Ширина столбцов
        col_letter = cell.column_letter
        ws.column_dimensions[col_letter].width = 18

    # Заголовок "Бренд" в столбце T (20)
    brand_cell = ws.cell(1, 20)
    brand_cell.value = "Бренд"
    brand_cell.fill = PatternFill(start_color="2D3748", end_color="2D3748", fill_type="solid")
    brand_cell.font = Font(bold=True, color="FFFFFF", size=12, name="Calibri")
    brand_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    brand_cell.border = calc_border
    ws.column_dimensions['T'].width = 18

    # ═══════════════════════════════════════════════════════════
    # 📋 ВЫПАДАЮЩИЙ СПИСОК КАТЕГОРИЙ для столбца F
    # ═══════════════════════════════════════════════════════════

    # Список категорий товаров
    categories = [
        "Очки",
        "Ракетка",
        "Кроссовки",
        "Куртка",
        "Штаны",
        "Шлем",
        "Ботинки борд",
        "Термо",
        "Очки для снега"
    ]

    # Создаем выпадающий список для столбца F (Категория товара)
    categories_formula = f'"{",".join(categories)}"'
    dv_category = DataValidation(
        type="list",
        formula1=categories_formula,
        allow_blank=True,
        showDropDown=False,  # False = показывать стрелку выпадающего списка
        showInputMessage=False,  # Не показывать примечание
        showErrorMessage=True
    )
    dv_category.error = "Выберите категорию из списка допустимых значений!"
    dv_category.errorTitle = "❌ Неверная категория"

    ws.add_data_validation(dv_category)
    # Применяем к столбцу F со строки 2 до 10000
    dv_category.add('F2:F10000')

    # ═══════════════════════════════════════════════════════════
    # ⚙️ ЛИСТ НАСТРОЕК
    # ═══════════════════════════════════════════════════════════

    settings_ws = wb.create_sheet("⚙️ Настройки")

    # Заголовок настроек
    settings_ws['A1'] = "⚙️ НАСТРОЙКИ РАСЧЕТОВ"
    settings_ws['A1'].font = Font(bold=True, size=16, name="Calibri")
    settings_ws['A1'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    settings_ws['A1'].font = Font(bold=True, color="FFFFFF", size=16, name="Calibri")
    settings_ws.merge_cells('A1:C1')

    # Курс валюты
    settings_ws['A3'] = "Курс EUR/RUB:"
    settings_ws['B3'] = 100.0  # Значение по умолчанию
    settings_ws['A3'].font = Font(bold=True, size=12)
    settings_ws['B3'].font = Font(size=12)
    settings_ws['B3'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    settings_ws['A4'] = "Надбавка:"
    settings_ws['B4'] = 0.5  # Значение по умолчанию
    settings_ws['A4'].font = Font(bold=True, size=12)
    settings_ws['B4'].font = Font(size=12)
    settings_ws['B4'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    settings_ws['A5'] = "Итоговый курс:"
    settings_ws['B5'] = "=B3+B4"
    settings_ws['A5'].font = Font(bold=True, size=12)
    settings_ws['B5'].font = Font(bold=True, size=14)
    settings_ws['B5'].fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

    # Таблица доставки
    settings_ws['A7'] = "📦 СТОИМОСТЬ ДОСТАВКИ (€)"
    settings_ws['A7'].font = Font(bold=True, size=14)
    settings_ws['A7'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    settings_ws['A7'].font = Font(bold=True, color="FFFFFF", size=14)
    settings_ws.merge_cells('A7:B7')

    settings_ws['A8'] = "Категория"
    settings_ws['B8'] = "Доставка (€)"
    settings_ws['A8'].font = Font(bold=True, size=11)
    settings_ws['B8'].font = Font(bold=True, size=11)
    settings_ws['A8'].fill = PatternFill(start_color="D0D0D0", end_color="D0D0D0", fill_type="solid")
    settings_ws['B8'].fill = PatternFill(start_color="D0D0D0", end_color="D0D0D0", fill_type="solid")

    # Таблица категорий и доставки
    delivery_table = [
        ("Очки", 12),
        ("Ракетка", 17),
        ("Кроссовки", 28),
        ("Куртка", 17),
        ("Штаны", 17),
        ("Шлем", 28),
        ("Ботинки борд", 25),
        ("Термо", 17),
        ("Очки для снега", 17)
    ]

    for idx, (cat, delivery) in enumerate(delivery_table, start=9):
        settings_ws[f'A{idx}'] = cat
        settings_ws[f'B{idx}'] = delivery
        settings_ws[f'A{idx}'].border = calc_border
        settings_ws[f'B{idx}'].border = calc_border

    # Секция БРЕНДЫ (столбец D)
    settings_ws['D1'] = "🏷️ БРЕНДЫ"
    settings_ws['D1'].font = Font(bold=True, color="FFFFFF", size=14, name="Calibri")
    settings_ws['D1'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

    settings_ws['D2'] = "Бренд"
    settings_ws['D2'].font = Font(bold=True, size=11)
    settings_ws['D2'].fill = PatternFill(start_color="D0D0D0", end_color="D0D0D0", fill_type="solid")

    brands_list = brands or [
        "Asics", "Adidas", "Bullpadel", "Drop Shot", "Head",
        "Joma", "Mizuno", "Nike", "Nox", "Oakley", "Puma", "Siux", "Wilson"
    ]
    for idx, brand in enumerate(brands_list, start=3):
        settings_ws[f'D{idx}'] = brand
        settings_ws[f'D{idx}'].border = calc_border

    # Ширина столбцов настроек
    settings_ws.column_dimensions['A'].width = 25
    settings_ws.column_dimensions['B'].width = 20
    settings_ws.column_dimensions['C'].width = 15
    settings_ws.column_dimensions['D'].width = 20

    wb.save(file_path)
    return file_path


# ═══════════════════════════════════════════════════════════
# 🕷️ ПАРСИНГ (из update_products.py)
# ═══════════════════════════════════════════════════════════

def get_images_dir(script_dir):
    """
    Определяет путь к папке images в зависимости от окружения.

    Приоритет:
    1. /data/images/ (если существует и НЕ пустая) - постоянное хранилище Amvera
    2. script_dir/images/ - из репозитория (fallback)
    """
    # Проверяем /data/images/ на Amvera
    data_path = Path('/data')
    if data_path.exists() and data_path.is_dir():
        data_images_dir = data_path / 'images'
        data_images_dir.mkdir(exist_ok=True)

        # Если там уже есть файлы - используем её
        if any(data_images_dir.iterdir()):
            return data_images_dir

    # Fallback: локальная папка или images из репозитория
    images_dir = script_dir / 'images'
    images_dir.mkdir(exist_ok=True)
    return images_dir


def clean_product_name(name):
    """Убирает русские (кириллические) слова из названия товара.

    Пример:
        "Bullpadel ракетка для паделя Vertex 04 2025" → "Bullpadel Vertex 04 2025"
    """
    if not name:
        return name

    import re

    # Разбиваем на слова
    words = name.split()

    # Оставляем только слова, которые не содержат кириллицу
    clean_words = []
    for word in words:
        # Проверяем, есть ли в слове хоть одна кириллическая буква
        if not re.search(r'[а-яА-ЯёЁ]', word):
            clean_words.append(word)

    # Собираем обратно в строку
    result = ' '.join(clean_words)

    # Убираем множественные пробелы
    result = re.sub(r'\s+', ' ', result).strip()

    return result


def download_image(image_url, save_dir, product_id):
    """Скачивает изображение и сохраняет локально (всегда перезаписывает)."""
    try:
        # Удаляем старые файлы для этого product_id (если есть)
        for old_file in save_dir.glob(f"product_{product_id}.*"):
            old_file.unlink()

        # Скачиваем изображение
        response = requests.get(image_url, timeout=10, stream=True)
        response.raise_for_status()

        # Определяем расширение файла
        content_type = response.headers.get('content-type', '')
        ext = '.jpg'
        if 'png' in content_type:
            ext = '.png'
        elif 'webp' in content_type:
            ext = '.webp'
        elif 'jpeg' in content_type or 'jpg' in content_type:
            ext = '.jpg'

        # Генерируем имя файла: product_1.webp, product_2.jpg и т.д.
        filename = f"product_{product_id}{ext}"
        filepath = save_dir / filename

        # Сохраняем файл (перезаписываем если существует)
        with open(filepath, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)

        return str(filepath.relative_to(save_dir.parent))
    except Exception as e:
        print(f"      ⚠️ Ошибка скачивания фото: {e}")
        return None


def parse_tradeinn_product(url, script_dir, product_id):
    """Парсит товар с tradeinn.com через HTML."""
    try:
        if '?' in url:
            url = url.split('?')[0]

        if '/en/' in url:
            url = url.replace('/en/', '/ru/')

        # Извлекаем product_id из URL (число перед /p)
        url_product_id = None
        product_id_match = re.search(r'/(\d+)/p/?$', url)
        if product_id_match:
            url_product_id = product_id_match.group(1)

        response = requests.get(url, timeout=10)
        response.raise_for_status()
        html = response.text

        name_match = re.search(r'<h1[^>]*>([^<]+)</h1>', html, re.IGNORECASE)
        name = name_match.group(1).strip() if name_match else "Без названия"
        # Убираем русские слова из названия
        name = clean_product_name(name)

        price_match = re.search(r'data-price="([^"]+)"', html, re.IGNORECASE)
        if price_match:
            try:
                price = float(price_match.group(1))
            except:
                price = 0
        else:
            price = 0

        # Парсим все фотки
        image_urls = []

        # МЕТОД 1: Ищем галерею с data-fancybox="gallery" - самый надёжный способ!
        # Эти ссылки ведут на полноразмерные изображения для просмотра
        gallery_links = re.findall(r'data-fancybox="gallery"[^>]*href="([^"]+)"', html, re.IGNORECASE)
        if gallery_links:
            for link in gallery_links:
                # Относительные ссылки преобразуем в абсолютные
                if link.startswith('/'):
                    link = 'https://www.tradeinn.com' + link
                if link not in image_urls:
                    image_urls.append(link)

        # МЕТОД 2: Если галерея не найдена, ищем через паттерн /f/категория/товар_X/
        # Учитываем суффиксы _2, _3, _4 и т.д. для разных фотографий
        if not image_urls and url_product_id:
            # Извлекаем ID категории из URL
            category_match = re.search(r'/(\d+)/\d+/p', url)
            if category_match:
                category_id = category_match.group(1)

                # Ищем все изображения с учётом суффиксов _2, _3, _4...
                # Паттерн: /f/14160/141608258_2/filename.webp
                pattern = rf'/f/{category_id}/{url_product_id}(?:_\d+)?/[^"\']+\.(?:jpg|jpeg|png|webp)'
                found_images = re.findall(pattern, html, re.IGNORECASE)

                for img in found_images:
                    full_url = 'https://www.tradeinn.com' + img if img.startswith('/') else img
                    if full_url not in image_urls:
                        image_urls.append(full_url)

        # МЕТОД 3: Ищем JSON объект с данными товара
        if not image_urls:
            json_match = re.search(r'var\s+product\s*=\s*(\{[^}]+images[^}]+\})', html, re.DOTALL)
            if not json_match:
                json_match = re.search(r'"images"\s*:\s*(\[[^\]]+\])', html, re.DOTALL)

            if json_match:
                try:
                    import json
                    images_data = json.loads(json_match.group(1))
                    if isinstance(images_data, list):
                        for img in images_data:
                            if isinstance(img, str) and 'tradeinn.com/f/' in img:
                                image_urls.append(img)
                    elif isinstance(images_data, dict) and 'images' in images_data:
                        for img in images_data['images']:
                            if isinstance(img, str) and 'tradeinn.com/f/' in img:
                                image_urls.append(img)
                except:
                    pass

        # МЕТОД 4: Широкий поиск всех изображений
        if not image_urls:
            all_images = re.findall(r'https://[^"\']+/f/\d+/\d+(?:_\d+)?/[^"\']+\.(?:jpg|jpeg|png|webp)', html, re.IGNORECASE)
            for img_url in all_images:
                if img_url not in image_urls and not any(x in img_url.lower() for x in ['_thumb', '_small', '_icon', 'logo']):
                    image_urls.append(img_url)

        # МЕТОД 5: Open Graph как запасной вариант
        if not image_urls:
            og_image = re.search(r'<meta property="og:image" content="([^"]+)"', html)
            if og_image and og_image.group(1).startswith('http'):
                image_urls.append(og_image.group(1))

        # Парсим размеры (для обуви, одежды)
        sizes = []

        # МЕТОД 1: JSON-LD структурированные данные (самый надежный для TradeInn!)
        json_ld_pattern = r'<script type="application/ld\+json">(.*?)</script>'
        json_ld_matches = re.findall(json_ld_pattern, html, re.DOTALL)

        for json_str in json_ld_matches:
            try:
                import json
                data = json.loads(json_str)

                # Ищем варианты товара (hasVariant)
                if isinstance(data, dict) and data.get('@type') == 'Product':
                    variants = data.get('hasVariant', [])
                    if variants:
                        for variant in variants:
                            # Извлекаем размер из имени варианта
                            variant_name = variant.get('name', '')
                            # Пример: "EU 42 1/2" или "EU 44"
                            size_match = re.search(r'EU\s+(\d+(?:\s*1/2)?)', variant_name)
                            if size_match:
                                size = size_match.group(1).strip()
                                if size not in sizes:
                                    sizes.append(size)
            except:
                pass

        # МЕТОД 2: Ищем размеры в select элементе
        if not sizes:
            size_select_match = re.findall(r'<option[^>]*value="size:([^"]+)"[^>]*>([^<]+)</option>', html, re.IGNORECASE)
            if size_select_match:
                for size_value, size_label in size_select_match:
                    size_clean = size_label.strip()
                    if size_clean and size_clean.lower() not in ['выберите размер', 'choose size', 'select']:
                        sizes.append(size_clean)

        # МЕТОД 3: Ищем размеры в data-атрибутах
        if not sizes:
            size_data_match = re.findall(r'data-size="([^"]+)"', html, re.IGNORECASE)
            if size_data_match:
                for size in size_data_match:
                    size_clean = size.strip()
                    if size_clean and len(size_clean) <= 10:
                        sizes.append(size_clean)

        # МЕТОД 4: Ищем текстовые паттерны "EU 42", "Size 42" (РАБОТАЕТ ДЛЯ TRADEINN!)
        if not sizes:
            text_patterns = [
                r'(?:EU|Size|Размер)\s+(\d{2}(?:\s*1/2)?)',  # EU 42, EU 42 1/2
                r'size["\']?\s*:\s*["\'](\d{2}(?:\s*1/2)?)["\']',  # JSON: "size":"42"
            ]

            for pattern in text_patterns:
                matches = re.findall(pattern, html, re.IGNORECASE)
                if matches:
                    for match in set(matches):
                        if match not in sizes:
                            sizes.append(match)

        # Убираем дубликаты и сортируем
        if sizes:
            sizes = list(dict.fromkeys(sizes))  # Убираем дубликаты, сохраняя порядок
            # Пытаемся отсортировать численно (учитываем дроби типа "42 1/2")
            def parse_size(s):
                # Преобразуем "42 1/2" в 42.5
                if '1/2' in s:
                    base = float(s.replace('1/2', '').strip())
                    return base + 0.5
                try:
                    return float(s.replace(',', '.'))
                except:
                    return 999

            try:
                sizes = sorted(set(sizes), key=parse_size)
            except:
                pass

        # Скачиваем только ПЕРВУЮ фотку (экономим место и трафик)
        images_dir = get_images_dir(script_dir)

        local_images = []
        if image_urls:
            # Берем только первую фотографию
            local_path = download_image(image_urls[0], images_dir, product_id)
            if local_path:
                local_images.append(local_path)

        return {
            "name": name,
            "price": price,
            "image_urls": ", ".join(image_urls) if image_urls else "",
            "local_images": ", ".join(local_images) if local_images else "",
            "sizes": ", ".join(sizes) if sizes else ""
        }, None

    except Exception as e:
        return None, f"Ошибка: {str(e)}"


def parse_generic_product(url, script_dir, product_id):
    """Универсальный парсер для других сайтов."""
    try:
        if '?' in url:
            url = url.split('?')[0]

        # Создаем session для сохранения cookies между запросами
        session = requests.Session()

        # Настройка headers для имитации реального браузера
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
        }
        session.headers.update(headers)

        # Для TradeInn устанавливаем страну доставки Armenia (id_pais=9)
        if 'tradeinn.com' in url:
            # Делаем запрос для установки страны Armenia через специальный endpoint
            try:
                # id_tienda=27 - это volleyball магазин, id_pais=9 - это Armenia
                country_setup_url = "https://www.tradeinn.com/get_dades.php?id_tienda=27&idioma=rus&id_pais=9&country_code_url="
                session.get(country_setup_url, timeout=5)
            except:
                pass  # Если не получилось - продолжаем без установки страны

        # Основной запрос страницы
        response = session.get(url, timeout=10)
        response.raise_for_status()
        html = response.text

        image_urls = []
        json_ld_match = re.search(r'<script type="application/ld\+json">(.*?)</script>', html, re.DOTALL)

        if json_ld_match:
            try:
                import json
                data = json.loads(json_ld_match.group(1))

                if data.get("@type") == "Product":
                    name = data.get("name", "Без названия")
                    description = data.get("description", "")[:100]

                    offers = data.get("offers", {})
                    if isinstance(offers, list):
                        offers = offers[0] if offers else {}

                    price_str = offers.get("price", "0")
                    try:
                        price = float(price_str)
                    except:
                        price = 0

                    # Парсим фотки из JSON-LD
                    images = data.get("image", [])
                    if isinstance(images, str):
                        images = [images]
                    elif isinstance(images, dict):
                        images = [images.get("url", "")]

                    for img in images:
                        if isinstance(img, str) and img.startswith('http'):
                            image_urls.append(img)
                        elif isinstance(img, dict) and img.get("url"):
                            image_urls.append(img["url"])

                    # Скачиваем только ПЕРВУЮ фотку (экономим место и трафик)
                    images_dir = script_dir / "images"
                    images_dir.mkdir(exist_ok=True)

                    local_images = []
                    if image_urls:
                        # Берем только первую фотографию
                        local_path = download_image(image_urls[0], images_dir, product_id)
                        if local_path:
                            local_images.append(local_path)

                    return {
                        "name": name,
                        "description": description,
                        "price": price,
                        "image_urls": ", ".join(image_urls) if image_urls else "",
                        "local_images": ", ".join(local_images) if local_images else ""
                    }, None
            except:
                pass

        name_match = re.search(r'<meta property="og:title" content="([^"]+)"', html)
        price_match = re.search(r'<meta property="product:price:amount" content="([^"]+)"', html)

        name = name_match.group(1) if name_match else "Без названия"
        # Убираем русские слова из названия
        name = clean_product_name(name)

        price = 0
        if price_match:
            try:
                price = float(price_match.group(1))
            except:
                price = 0

        # Парсим фотки через Open Graph
        og_images = re.findall(r'<meta property="og:image" content="([^"]+)"', html)
        for img in og_images:
            if img.startswith('http'):
                image_urls.append(img)

        # Скачиваем только ПЕРВУЮ фотку (экономим место и трафик)
        images_dir = get_images_dir(script_dir)

        local_images = []
        if image_urls:
            # Берем только первую фотографию
            local_path = download_image(image_urls[0], images_dir, product_id)
            if local_path:
                local_images.append(local_path)

        return {
            "name": name,
            "price": price,
            "image_urls": ", ".join(image_urls) if image_urls else "",
            "local_images": ", ".join(local_images) if local_images else ""
        }, None

    except Exception as e:
        return None, f"Ошибка: {str(e)}"


def parse_product(url, script_dir, product_id):
    """Определяет сайт и парсит товар."""
    if not url or not url.startswith("http"):
        return None, "Некорректный URL"

    if "tradeinn.com" in url:
        return parse_tradeinn_product(url, script_dir, product_id)
    else:
        return parse_generic_product(url, script_dir, product_id)


# ═══════════════════════════════════════════════════════════
# 🖥️ GUI ПРИЛОЖЕНИЕ
# ═══════════════════════════════════════════════════════════

class ParserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🕷️ Парсер товаров для Telegram магазина")
        self.root.geometry("900x650")
        self.root.resizable(True, True)

        # Папка скрипта
        self.script_dir = Path(__file__).parent
        self.file_path = self.script_dir / "products_links.xlsx"
        self.settings_file = self.script_dir / "parser_settings.json"

        # Стили
        style = ttk.Style()
        style.theme_use('clam')

        # Главный фрейм
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок приложения
        title_label = tk.Label(
            main_frame,
            text="🕷️ Парсер товаров для Telegram магазина",
            font=("Segoe UI", 16, "bold"),
            fg="#1F4E78"
        )
        title_label.pack(pady=(0, 20))

        # ═══════════════════════════════════════════════════════════
        # 📑 ВКЛАДКИ (Notebook)
        # ═══════════════════════════════════════════════════════════

        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        # ═══════════════════════════════════════════════════════════
        # 📑 ВКЛАДКА 1: ПАРСИНГ ТОВАРОВ
        # ═══════════════════════════════════════════════════════════

        tab_parser = ttk.Frame(notebook, padding="10")
        notebook.add(tab_parser, text="🕷️ Парсинг товаров")

        # Фрейм для кнопок
        button_frame = ttk.Frame(tab_parser)
        button_frame.pack(pady=10)

        # Кнопка создания шаблона
        self.create_btn = tk.Button(
            button_frame,
            text="📄 Создать шаблонный файл",
            command=self.create_template_clicked,
            bg="#4CAF50",
            fg="white",
            font=("Segoe UI", 12, "bold"),
            padx=20,
            pady=10,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2
        )
        self.create_btn.pack(side=tk.LEFT, padx=10)

        # Кнопка парсинга
        self.parse_btn = tk.Button(
            button_frame,
            text="🚀 Спарсить из Excel файла",
            command=self.parse_clicked,
            bg="#2196F3",
            fg="white",
            font=("Segoe UI", 12, "bold"),
            padx=20,
            pady=10,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2
        )
        self.parse_btn.pack(side=tk.LEFT, padx=10)

        # Кнопка архивирования
        self.archive_btn = tk.Button(
            button_frame,
            text="📦 Заархивировать для бота",
            command=self.archive_clicked,
            bg="#FF9800",
            fg="white",
            font=("Segoe UI", 12, "bold"),
            padx=20,
            pady=10,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2
        )
        self.archive_btn.pack(side=tk.LEFT, padx=10)

        # Выбор количества потоков
        threads_frame = tk.Frame(tab_parser, bg="#f5f5f5")
        threads_frame.pack(pady=5)

        tk.Label(
            threads_frame,
            text="⚡ Потоки:",
            font=("Segoe UI", 11, "bold"),
            bg="#f5f5f5",
            fg="#333"
        ).pack(side=tk.LEFT, padx=(0, 8))

        self.threads_var = tk.IntVar(value=saved.get("threads", 5))
        self.threads_spinbox = tk.Spinbox(
            threads_frame,
            from_=1,
            to=10,
            textvariable=self.threads_var,
            width=3,
            font=("Segoe UI", 12, "bold"),
            justify=tk.CENTER,
            state="readonly"
        )
        self.threads_spinbox.pack(side=tk.LEFT)

        tk.Label(
            threads_frame,
            text="(1 = медленно, 5 = оптимально, 10 = максимум)",
            font=("Segoe UI", 9),
            bg="#f5f5f5",
            fg="#888"
        ).pack(side=tk.LEFT, padx=(8, 0))

        # Информационная панель
        info_frame = ttk.LabelFrame(tab_parser, text="ℹ️ Информация", padding="10")
        info_frame.pack(fill=tk.X, pady=10)

        self.info_label = tk.Label(
            info_frame,
            text=f"📁 Файл: {self.file_path.name}\n📂 Папка: {self.script_dir}",
            font=("Segoe UI", 10),
            justify=tk.LEFT,
            fg="#555"
        )
        self.info_label.pack(anchor=tk.W)

        # Лог-панель
        log_frame = ttk.LabelFrame(tab_parser, text="📋 Журнал работы", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Текстовое поле для логов
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            font=("Consolas", 10),
            bg="#F5F5F5",
            fg="#333",
            height=15
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Статус бар
        self.status_bar = tk.Label(
            root,
            text="Готов к работе",
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            bg="#1F4E78",
            fg="white",
            font=("Segoe UI", 10),
            padx=10,
            pady=5
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Приветственное сообщение
        self.log("=" * 80)
        self.log("🎉 Добро пожаловать в Парсер товаров!")
        self.log("=" * 80)

        # ═══════════════════════════════════════════════════════════
        # 📑 ВКЛАДКА 2: КУРС ВАЛЮТЫ
        # ═══════════════════════════════════════════════════════════

        tab_currency = ttk.Frame(notebook, padding="0")
        notebook.add(tab_currency, text="💱 Курс валюты")

        # Создаем Canvas с прокруткой для всего содержимого
        currency_canvas = tk.Canvas(tab_currency, highlightthickness=0)
        currency_scrollbar = ttk.Scrollbar(tab_currency, orient="vertical", command=currency_canvas.yview)
        currency_scrollable_frame = ttk.Frame(currency_canvas, padding="20")

        currency_scrollable_frame.bind(
            "<Configure>",
            lambda _: currency_canvas.configure(scrollregion=currency_canvas.bbox("all"))
        )

        currency_canvas.create_window((0, 0), window=currency_scrollable_frame, anchor="nw")
        currency_canvas.configure(yscrollcommand=currency_scrollbar.set)

        currency_canvas.pack(side="left", fill="both", expand=True)
        currency_scrollbar.pack(side="right", fill="y")

        # Привязываем прокрутку мышью
        def _on_mousewheel(event):
            currency_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        currency_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Заголовок вкладки
        currency_title = tk.Label(
            currency_scrollable_frame,
            text="💱 Управление курсом валюты EUR → RUB",
            font=("Segoe UI", 14, "bold"),
            fg="#1F4E78"
        )
        currency_title.pack(pady=(0, 15))

        # Фрейм для текущего курса и настроек (компактно)
        current_settings_frame = ttk.LabelFrame(currency_scrollable_frame, text="📊 Курс и настройки", padding="15")
        current_settings_frame.pack(fill=tk.X, pady=10)

        self.currency_rate_label = tk.Label(
            current_settings_frame,
            text="Курс EUR/RUB: загрузка...",
            font=("Segoe UI", 11, "bold"),
            fg="#2196F3"
        )
        self.currency_rate_label.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=5)

        self.last_update_label = tk.Label(
            current_settings_frame,
            text="Последнее обновление: -",
            font=("Segoe UI", 9),
            fg="#666"
        )
        self.last_update_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))

        # Надбавка к курсу
        markup_label = tk.Label(
            current_settings_frame,
            text="Надбавка к курсу (+):",
            font=("Segoe UI", 10)
        )
        markup_label.grid(row=2, column=0, sticky=tk.W, pady=5, padx=(0, 10))

        self.markup_entry = tk.Entry(
            current_settings_frame,
            font=("Segoe UI", 10),
            width=10
        )
        self.markup_entry.insert(0, str(saved.get("markup", 0.5)))
        self.markup_entry.grid(row=2, column=1, sticky=tk.W, pady=5)

        markup_hint = tk.Label(
            current_settings_frame,
            text="₽",
            font=("Segoe UI", 10),
            fg="#666"
        )
        markup_hint.grid(row=2, column=2, sticky=tk.W, pady=5, padx=(5, 0))

        # ═══════════════════════════════════════════════════════════
        # 📋 КАТЕГОРИИ ТОВАРОВ И СТОИМОСТЬ ДОСТАВКИ
        # ═══════════════════════════════════════════════════════════

        # Фрейм для категорий
        self.categories_main_frame = ttk.LabelFrame(currency_scrollable_frame, text="📋 Категории и стоимость доставки (€)", padding="15")
        self.categories_main_frame.pack(fill=tk.X, pady=10)

        # Загружаем сохранённые настройки (или дефолтные)
        saved = self._load_settings()

        self.categories_data = saved.get("categories", [
            {"name": "Очки", "delivery": 12},
            {"name": "Ракетка", "delivery": 17},
            {"name": "Кроссовки", "delivery": 28},
            {"name": "Куртка", "delivery": 17},
            {"name": "Штаны", "delivery": 17},
            {"name": "Шлем", "delivery": 28},
            {"name": "Ботинки борд", "delivery": 25},
            {"name": "Термо", "delivery": 17},
            {"name": "Очки для снега", "delivery": 17}
        ])

        # Создаем фрейм для таблицы (будет перерисовываться)
        self.categories_table_frame = tk.Frame(self.categories_main_frame)
        self.categories_table_frame.pack(fill=tk.X)

        # Кнопки управления
        buttons_frame = tk.Frame(self.categories_main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)

        add_category_btn = tk.Button(
            buttons_frame,
            text="➕ Добавить категорию",
            command=self.add_category_dialog,
            bg="#2196F3",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=5,
            cursor="hand2"
        )
        add_category_btn.pack(side=tk.LEFT, padx=5)

        save_categories_btn = tk.Button(
            buttons_frame,
            text="💾 Сохранить изменения",
            command=self.save_category_changes,
            bg="#4CAF50",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=5,
            cursor="hand2"
        )
        save_categories_btn.pack(side=tk.LEFT, padx=5)

        # Отрисовываем таблицу категорий
        self.refresh_categories_table()

        # ═══════════════════════════════════════════════════════════
        # 🏷️ БРЕНДЫ
        # ═══════════════════════════════════════════════════════════

        self.brands_data = saved.get("brands", [
            "Asics", "Adidas", "Bullpadel", "Drop Shot", "Head",
            "Joma", "Mizuno", "Nike", "Nox", "Oakley", "Puma", "Siux", "Wilson"
        ])

        self.brands_main_frame = ttk.LabelFrame(currency_scrollable_frame, text="🏷️ Бренды (для авто-определения из названий)", padding="15")
        self.brands_main_frame.pack(fill=tk.X, pady=10)

        self.brands_table_frame = tk.Frame(self.brands_main_frame)
        self.brands_table_frame.pack(fill=tk.X)

        brands_buttons_frame = tk.Frame(self.brands_main_frame)
        brands_buttons_frame.pack(fill=tk.X, pady=10)

        add_brand_btn = tk.Button(
            brands_buttons_frame,
            text="➕ Добавить бренд",
            command=self.add_brand_dialog,
            bg="#2196F3",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=5,
            cursor="hand2"
        )
        add_brand_btn.pack(side=tk.LEFT, padx=5)

        self.refresh_brands_table()

        # ═══════════════════════════════════════════════════════════
        # 📊 КОЭФФИЦИЕНТЫ НАЦЕНКИ
        # ═══════════════════════════════════════════════════════════

        coef_frame = ttk.LabelFrame(currency_scrollable_frame, text="📊 Коэффициенты наценки (%)", padding="15")
        coef_frame.pack(fill=tk.X, pady=10)

        # Кэф Пети
        tk.Label(coef_frame, text="📊 Кэф Пети (на основе Закупки)", font=("Segoe UI", 10, "bold"), fg="#4CAF50").grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        tk.Label(coef_frame, text="Закупка < 15,000₽:", font=("Segoe UI", 9)).grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        tk.Label(coef_frame, text="10%", font=("Segoe UI", 9, "bold")).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)

        tk.Label(coef_frame, text="Закупка ≤ 30,000₽:", font=("Segoe UI", 9)).grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        tk.Label(coef_frame, text="9%", font=("Segoe UI", 9, "bold")).grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)

        tk.Label(coef_frame, text="Закупка > 30,000₽:", font=("Segoe UI", 9)).grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        tk.Label(coef_frame, text="8%", font=("Segoe UI", 9, "bold")).grid(row=3, column=1, sticky=tk.W, padx=5, pady=2)

        # Разделитель
        ttk.Separator(coef_frame, orient="horizontal").grid(row=4, column=0, columnspan=3, sticky="ew", pady=10)

        # Наш Кэф
        tk.Label(coef_frame, text="💰 Наш Кэф (на основе Закупки)", font=("Segoe UI", 10, "bold"), fg="#2196F3").grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        tk.Label(coef_frame, text="Закупка < 10,000₽:", font=("Segoe UI", 9)).grid(row=6, column=0, sticky=tk.W, padx=5, pady=2)
        tk.Label(coef_frame, text="17%", font=("Segoe UI", 9, "bold")).grid(row=6, column=1, sticky=tk.W, padx=5, pady=2)

        tk.Label(coef_frame, text="Закупка ≤ 20,000₽:", font=("Segoe UI", 9)).grid(row=7, column=0, sticky=tk.W, padx=5, pady=2)
        tk.Label(coef_frame, text="15%", font=("Segoe UI", 9, "bold")).grid(row=7, column=1, sticky=tk.W, padx=5, pady=2)

        tk.Label(coef_frame, text="Закупка ≤ 30,000₽:", font=("Segoe UI", 9)).grid(row=8, column=0, sticky=tk.W, padx=5, pady=2)
        tk.Label(coef_frame, text="14%", font=("Segoe UI", 9, "bold")).grid(row=8, column=1, sticky=tk.W, padx=5, pady=2)

        tk.Label(coef_frame, text="Закупка > 30,000₽:", font=("Segoe UI", 9)).grid(row=9, column=0, sticky=tk.W, padx=5, pady=2)
        tk.Label(coef_frame, text="13%", font=("Segoe UI", 9, "bold")).grid(row=9, column=1, sticky=tk.W, padx=5, pady=2)

        # ═══════════════════════════════════════════════════════════
        # 📐 ФОРМУЛЫ РАСЧЕТА (краткий справочник)
        # ═══════════════════════════════════════════════════════════

        formulas_frame = ttk.LabelFrame(currency_scrollable_frame, text="📐 Excel формулы (краткий справочник)", padding="15")
        formulas_frame.pack(fill=tk.X, pady=10)

        formulas_text = tk.Label(
            formulas_frame,
            text=(
                "L (Доставка₽)         = VLOOKUP(Категория, Таблица_доставки) × Курс\n"
                "M (Закупка₽)          = Доставка + (Цена€ × Курс)\n"
                "N (Кэф Пети %)        = IFS(Закупка<15000, 10%, Закупка≤30000, 9%, Закупка>30000, 8%)\n"
                "O (Наш Кэф %)         = IFS(Закупка<10000, 17%, Закупка≤20000, 15%, Закупка≤30000, 14%, Закупка>30000, 13%)\n"
                "P (Цена с дост.₽)     = Закупка × (1 + Кэф_Пети + Наш_Кэф)\n"
                "Q (Цена без дост.₽)   = Цена_с_доставкой - Доставка\n"
                "R (Наша Маржа₽)       = Закупка × Наш_Кэф\n"
                "S (Маржа Пети₽)       = Закупка × Кэф_Пети"
            ),
            font=("Consolas", 9),
            fg="#333",
            justify=tk.LEFT
        )
        formulas_text.pack(pady=5)

        # Пояснение
        formulas_hint = tk.Label(
            formulas_frame,
            text="💡 Все формулы автоматически вставляются в Excel при нажатии '📊 Применить формулы к Excel'",
            font=("Segoe UI", 9),
            fg="#666"
        )
        formulas_hint.pack(pady=(10, 0))

        # ═══════════════════════════════════════════════════════════
        # 🎯 КНОПКИ УПРАВЛЕНИЯ
        # ═══════════════════════════════════════════════════════════

        # Кнопки управления курсом и формулами
        currency_buttons_frame = ttk.Frame(currency_scrollable_frame)
        currency_buttons_frame.pack(pady=20)

        # Кнопка обновления курса
        self.update_rate_btn = tk.Button(
            currency_buttons_frame,
            text="🔄 Обновить курс",
            command=self.update_currency_rate,
            bg="#4CAF50",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2
        )
        self.update_rate_btn.pack(side=tk.LEFT, padx=10)

        # Кнопка применения формул к Excel
        self.apply_formulas_btn = tk.Button(
            currency_buttons_frame,
            text="📊 Применить формулы к Excel",
            command=self.apply_formulas_to_excel,
            bg="#2196F3",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2
        )
        self.apply_formulas_btn.pack(side=tk.LEFT, padx=10)

        # Инициализация: загружаем курс при запуске
        self.current_eur_rub = 0
        self.update_currency_rate()
        self.log("")
        self.log("📝 Инструкция:")
        self.log("1. Нажми '📄 Создать шаблонный файл' если файла ещё нет")
        self.log("2. Открой Excel файл и вставь ссылки на товары в колонку A")
        self.log("3. Заполни Группу, Подгруппу и Эмодзи (опционально)")
        self.log("4. Нажми '🚀 Спарсить из Excel файла'")
        self.log("")

        if self.file_path.exists():
            self.log(f"✅ Файл {self.file_path.name} уже существует")
            self.update_status("Файл найден, готов к парсингу")
        else:
            self.log(f"⚠️ Файл {self.file_path.name} не найден - создай шаблон")
            self.update_status("Создай шаблонный файл для начала работы")

    def log(self, message, color=None):
        """Добавляет сообщение в лог."""
        if color:
            tag = f"color_{color}"
            self.log_text.tag_configure(tag, foreground=color)
            self.log_text.insert(tk.END, message + "\n", tag)
        else:
            self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def update_status(self, message):
        """Обновляет статус бар."""
        self.status_bar.config(text=message)
        self.root.update()

    def create_template_clicked(self):
        """Обработчик кнопки создания шаблона."""
        self.update_status("Создание шаблона...")
        self.log("\n" + "=" * 80)
        self.log("📄 СОЗДАНИЕ ШАБЛОНА")
        self.log("=" * 80)

        try:
            if self.file_path.exists():
                response = messagebox.askyesno(
                    "Файл существует",
                    f"Файл {self.file_path.name} уже существует.\nПерезаписать?"
                )
                if not response:
                    self.log("❌ Отменено пользователем")
                    self.update_status("Отменено")
                    return

            file_path = create_beautiful_template(self.file_path, brands=self.get_brands_from_ui())
            self.log(f"✅ Шаблон создан: {file_path}")
            self.log("")
            self.log("📝 Следующие шаги:")
            self.log(f"   1. Открой файл: {file_path}")
            self.log("   2. Вставь ссылки в колонку 'URL'")
            self.log("   3. Заполни Группу, Подгруппу, Эмодзи")
            self.log("   4. Нажми 'Спарсить из Excel файла'")
            self.log("")

            self.update_status("✅ Шаблон успешно создан")

            messagebox.showinfo(
                "Успех",
                f"Шаблон создан!\n\n📁 {file_path}\n\nТеперь заполни ссылки и запусти парсинг."
            )

        except Exception as e:
            self.log(f"❌ Ошибка создания шаблона: {e}")
            self.update_status("❌ Ошибка создания шаблона")
            messagebox.showerror("Ошибка", f"Не удалось создать шаблон:\n{e}")

    def archive_clicked(self):
        """Обработчик кнопки архивирования."""
        self.update_status("📦 Создание архива...")
        self.log("\n" + "=" * 80)
        self.log("📦 СОЗДАНИЕ ZIP АРХИВА ДЛЯ БОТА")
        self.log("=" * 80)
        self.log("")

        try:
            # Проверяем наличие файлов
            if not self.file_path.exists():
                messagebox.showwarning(
                    "Файл не найден",
                    f"Excel файл {self.file_path.name} не найден!\n\nСначала спарси товары."
                )
                self.log("❌ Excel файл не найден")
                self.update_status("❌ Excel файл не найден")
                return

            images_dir = get_images_dir(self.script_dir)
            if not images_dir.exists() or not any(images_dir.iterdir()):
                messagebox.showwarning(
                    "Папка images пуста",
                    "Папка images/ не найдена или пуста!\n\nСначала спарси товары."
                )
                self.log("❌ Папка images/ пуста")
                self.update_status("❌ Папка images/ пуста")
                return

            # Создаём ZIP архив
            archive_name = f"catalog_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
            archive_path = self.script_dir / archive_name

            self.log(f"📦 Создаю архив: {archive_name}")
            self.log("")

            with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Добавляем Excel файл
                self.log(f"   ✅ Добавляю {self.file_path.name}")
                zipf.write(self.file_path, self.file_path.name)

                # Добавляем все фотографии из папки images
                image_count = 0
                for image_file in images_dir.iterdir():
                    if image_file.is_file():
                        # Добавляем с путём images/filename
                        arcname = f"images/{image_file.name}"
                        zipf.write(image_file, arcname)
                        image_count += 1

                self.log(f"   ✅ Добавлено фотографий: {image_count}")

            archive_size = archive_path.stat().st_size / 1024 / 1024  # MB

            self.log("")
            self.log("=" * 80)
            self.log(f"✅ Архив создан: {archive_name}")
            self.log(f"📁 Путь: {archive_path}")
            self.log(f"💾 Размер: {archive_size:.2f} MB")
            self.log("")
            self.log("📤 Следующие шаги:")
            self.log("   1. Отправь этот ZIP архив своему Telegram боту")
            self.log("   2. Бот автоматически распакует и обновит каталог")
            self.log("   3. Используй /shop чтобы открыть магазин с фотографиями")
            self.log("=" * 80)
            self.log("")

            self.update_status(f"✅ Архив создан: {archive_name}")

            messagebox.showinfo(
                "Архив создан!",
                f"✅ ZIP архив создан!\n\n"
                f"📁 {archive_name}\n"
                f"💾 Размер: {archive_size:.2f} MB\n\n"
                f"Отправь этот архив своему Telegram боту\n"
                f"для обновления каталога!"
            )

            # Открываем папку с архивом
            import subprocess
            subprocess.Popen(f'explorer /select,"{archive_path}"')

        except Exception as e:
            self.log(f"\n❌ ОШИБКА: {e}")
            self.update_status("❌ Ошибка создания архива")
            messagebox.showerror("Ошибка", f"Ошибка создания архива:\n{e}")

    def parse_clicked(self):
        """Обработчик кнопки парсинга."""
        if not self.file_path.exists():
            messagebox.showwarning(
                "Файл не найден",
                f"Файл {self.file_path.name} не найден!\n\nСначала создай шаблон."
            )
            return

        # Запускаем парсинг в отдельном потоке
        thread = threading.Thread(target=self.parse_excel, daemon=True)
        thread.start()

    def parse_excel(self):
        """Парсит товары из Excel (многопоточно)."""
        # Блокируем кнопки
        self.create_btn.config(state=tk.DISABLED)
        self.parse_btn.config(state=tk.DISABLED)

        num_threads = self.threads_var.get()
        self.update_status(f"🕷️ Парсинг в процессе ({num_threads} потоков)...")
        self.log("\n" + "=" * 80)
        self.log(f"🕷️ ПАРСИНГ ТОВАРОВ ({num_threads} потоков)")
        self.log("=" * 80)
        self.log("")

        try:
            wb = load_workbook(self.file_path)
            ws = wb.active

            updated_count = 0
            error_count = 0
            total_rows = ws.max_row - 1  # Минус заголовок

            # ═══════════════════════════════════════════════════════════
            # 1. Собираем задачи для парсинга
            # ═══════════════════════════════════════════════════════════
            tasks = []
            for row_num in range(2, ws.max_row + 1):
                url = ws.cell(row_num, 1).value
                if not url or not url.startswith("http"):
                    self.log(f"[{row_num - 1}/{total_rows}] ⏭️ Пропущено (нет URL)")
                    ws.cell(row_num, 11).value = "Пропущено (нет URL)"
                    continue
                tasks.append((row_num, url, row_num - 1))

            self.log(f"📋 Найдено {len(tasks)} ссылок для парсинга\n")

            # ═══════════════════════════════════════════════════════════
            # 2. Параллельный парсинг через ThreadPoolExecutor
            # ═══════════════════════════════════════════════════════════
            results = {}  # {row_num: (product_data, error)}
            completed = 0

            with ThreadPoolExecutor(max_workers=num_threads) as executor:
                future_to_row = {
                    executor.submit(parse_product, url, self.script_dir, pid): (row_num, url, pid)
                    for row_num, url, pid in tasks
                }

                for future in as_completed(future_to_row):
                    row_num, url, pid = future_to_row[future]
                    completed += 1

                    try:
                        product_data, error = future.result()
                        results[row_num] = (product_data, error)

                        if error:
                            self.log(f"[{completed}/{len(tasks)}] ❌ #{pid}: {error}", color="red")
                        else:
                            photos = len(product_data['image_urls'].split(',')) if product_data.get('image_urls') else 0
                            price = product_data.get('price')

                            if not price or not photos:
                                missing = []
                                if not price: missing.append("нет цены")
                                if not photos: missing.append("нет фото")
                                self.log(f"[{completed}/{len(tasks)}] ⚠️ #{pid}: {product_data['name']} | {price or '???'}€ | 📷{photos} — {', '.join(missing)}", color="red")
                            else:
                                self.log(f"[{completed}/{len(tasks)}] ✅ #{pid}: {product_data['name']} | {price}€ | 📷{photos}")
                    except Exception as e:
                        results[row_num] = (None, str(e))
                        self.log(f"[{completed}/{len(tasks)}] ❌ #{pid}: {e}", color="red")

                    self.update_status(f"🕷️ Парсинг: {completed}/{len(tasks)} ({num_threads} потоков)")

            # ═══════════════════════════════════════════════════════════
            # 3. Читаем список брендов из настроек
            # ═══════════════════════════════════════════════════════════
            brands_list = []
            if "⚙️ Настройки" in wb.sheetnames:
                settings_ws = wb["⚙️ Настройки"]
                for row in range(3, 100):
                    brand = settings_ws[f'D{row}'].value
                    if brand and str(brand).strip():
                        brands_list.append(str(brand).strip())
                    elif row > 10:
                        break
            self.log(f"\n🏷️ Брендов: {len(brands_list)} ({', '.join(brands_list[:5])}{'...' if len(brands_list) > 5 else ''})")

            # ═══════════════════════════════════════════════════════════
            # 4. Записываем результаты в Excel (последовательно)
            # ═══════════════════════════════════════════════════════════
            self.log("📝 Записываю результаты в Excel...")

            data_border = Border(
                left=Side(style='thin', color='D0D0D0'),
                right=Side(style='thin', color='D0D0D0'),
                top=Side(style='thin', color='D0D0D0'),
                bottom=Side(style='thin', color='D0D0D0')
            )
            data_font = Font(size=11, name="Calibri")
            left_alignment = Alignment(horizontal="left", vertical="center")
            center_alignment = Alignment(horizontal="center", vertical="center")

            for row_num, (product_data, error) in sorted(results.items()):
                if error:
                    ws.cell(row_num, 11).value = error
                    error_count += 1
                else:
                    # Проверяем, есть ли вручную заполненные данные
                    existing_category = ws.cell(row_num, 6).value
                    if existing_category:
                        self.log(f"   📋 #{row_num-1}: Категория сохранена: {existing_category}")

                    # Обновляем ТОЛЬКО автоматически заполняемые поля
                    ws.cell(row_num, 2).value = product_data['name']           # B: Название
                    ws.cell(row_num, 3).value = product_data['price']          # C: Цена
                    # D: Группа (НЕ ТРОГАЕМ - заполняется вручную)
                    # E: Подгруппа (НЕ ТРОГАЕМ - заполняется вручную)
                    # F: Категория товара (НЕ ТРОГАЕМ - заполняется вручную)
                    ws.cell(row_num, 7).value = product_data['image_urls']     # G: URL фото
                    ws.cell(row_num, 8).value = product_data['local_images']   # H: Локальное фото
                    ws.cell(row_num, 9).value = product_data.get('sizes', '')  # I: Размеры
                    ws.cell(row_num, 10).value = datetime.now().strftime("%Y-%m-%d %H:%M")  # J: Обновление
                    ws.cell(row_num, 11).value = "✅ Обновлено"                # K: Статус

                    # T(20): Бренд — определяем из названия
                    detected_brand = ""
                    name_lower = product_data['name'].lower()
                    for brand in brands_list:
                        if brand.lower() in name_lower:
                            detected_brand = brand
                            break
                    ws.cell(row_num, 20).value = detected_brand  # T: Бренд

                    # Применяем оформление к ячейкам
                    cell_a = ws.cell(row_num, 1)
                    cell_a.border = data_border
                    cell_a.font = data_font
                    cell_a.alignment = left_alignment

                    cell_b = ws.cell(row_num, 2)
                    cell_b.border = data_border
                    cell_b.font = data_font
                    cell_b.alignment = left_alignment
                    cell_b.fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")

                    cell_c = ws.cell(row_num, 3)
                    cell_c.border = data_border
                    cell_c.font = data_font
                    cell_c.alignment = center_alignment
                    cell_c.number_format = '#,##0.00'
                    cell_c.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

                    for col in [4, 5, 6]:
                        cell = ws.cell(row_num, col)
                        cell.border = data_border
                        cell.font = data_font
                        cell.alignment = left_alignment

                    cell_g = ws.cell(row_num, 7)
                    cell_g.border = data_border
                    cell_g.font = Font(size=9, name="Calibri")
                    cell_g.alignment = left_alignment

                    cell_h = ws.cell(row_num, 8)
                    cell_h.border = data_border
                    cell_h.font = data_font
                    cell_h.alignment = left_alignment

                    cell_i = ws.cell(row_num, 9)
                    cell_i.border = data_border
                    cell_i.font = data_font
                    cell_i.alignment = center_alignment

                    cell_j = ws.cell(row_num, 10)
                    cell_j.border = data_border
                    cell_j.font = Font(size=10, name="Calibri")
                    cell_j.alignment = center_alignment

                    cell_k = ws.cell(row_num, 11)
                    cell_k.border = data_border
                    cell_k.font = data_font
                    cell_k.alignment = center_alignment

                    for col in range(12, 20):
                        cell = ws.cell(row_num, col)
                        if cell.value is not None:
                            cell.border = data_border

                    updated_count += 1

            # ═══════════════════════════════════════════════════════════
            # 📋 ОБНОВЛЯЕМ ВЫПАДАЮЩИЙ СПИСОК КАТЕГОРИЙ
            # ═══════════════════════════════════════════════════════════

            # Удаляем старую валидацию (если есть)
            ws.data_validations.dataValidation = [
                dv for dv in ws.data_validations.dataValidation
                if dv.sqref and 'F' not in str(dv.sqref).split(':')[0]
            ]

            # Создаем новую валидацию с категориями из листа настроек (если есть)
            if "⚙️ Настройки" in wb.sheetnames:
                # Вычисляем конечную строку динамически на основе количества категорий
                settings_ws = wb["⚙️ Настройки"]
                last_row = 8  # Строка перед первой категорией

                # Ищем последнюю заполненную строку с категорией (начиная со строки 9)
                for row in range(9, 100):
                    if settings_ws[f'A{row}'].value:
                        last_row = row
                    else:
                        break

                # Используем ссылку на лист настроек для динамического списка
                categories_formula = f"'⚙️ Настройки'!$A$9:$A${last_row}"
                dv_category = DataValidation(
                    type="list",
                    formula1=categories_formula,
                    allow_blank=True,
                    showDropDown=False,  # False = показывать стрелку выпадающего списка
                    showInputMessage=False,  # Не показывать примечание
                    showErrorMessage=True
                )
            else:
                # Если листа настроек нет, используем статический список
                categories = ["Очки", "Ракетка", "Кроссовки", "Куртка", "Штаны", "Шлем", "Ботинки борд", "Термо", "Очки для снега"]
                categories_formula = f'"{",".join(categories)}"'
                dv_category = DataValidation(
                    type="list",
                    formula1=categories_formula,
                    allow_blank=True,
                    showDropDown=False,  # False = показывать стрелку выпадающего списка
                    showInputMessage=False,  # Не показывать примечание
                    showErrorMessage=True
                )

            dv_category.error = "Выберите категорию из списка допустимых значений!"
            dv_category.errorTitle = "❌ Неверная категория"

            ws.add_data_validation(dv_category)
            # Применяем к столбцу F для всех строк (включая новые)
            max_row = ws.max_row if ws.max_row > 2 else 10000
            dv_category.add(f'F2:F{max_row}')

            # Сохраняем
            wb.save(self.file_path)

            self.log("")
            self.log("=" * 80)
            self.log(f"✅ Обновлено товаров: {updated_count}")
            self.log(f"❌ Ошибок: {error_count}")
            self.log(f"📄 Файл сохранён: {self.file_path}")
            self.log(f"📁 Фотки сохранены в: {self.script_dir / 'images'}")
            self.log("=" * 80)
            self.log("")

            # ═══════════════════════════════════════════════════════════
            # 📊 АВТОМАТИЧЕСКОЕ ПРИМЕНЕНИЕ ФОРМУЛ ПОСЛЕ ПАРСИНГА
            # ═══════════════════════════════════════════════════════════

            self.log("📊 Применяю формулы расчета к товарам...")
            self.update_status("📊 Применение формул...")

            # Применяем формулы без messagebox (тихо)
            try:
                self.apply_formulas_silently()
                self.log("✅ Формулы применены! Столбцы L-S обновлены.")
            except Exception as e:
                self.log(f"⚠️ Не удалось применить формулы: {e}")
                self.log("   Можно применить формулы вручную на вкладке 'Курс валюты'")

            self.log("")
            self.update_status(f"✅ Парсинг завершён: {updated_count} товаров обновлено")

            messagebox.showinfo(
                "Парсинг завершён",
                f"✅ Обновлено товаров: {updated_count}\n❌ Ошибок: {error_count}\n\n📊 Формулы применены автоматически!\n\n📄 {self.file_path}\n📁 Фотки: {self.script_dir / 'images'}"
            )

        except Exception as e:
            self.log(f"\n❌ ОШИБКА: {e}")
            self.update_status("❌ Ошибка парсинга")
            messagebox.showerror("Ошибка", f"Ошибка парсинга:\n{e}")

        finally:
            # Разблокируем кнопки
            self.create_btn.config(state=tk.NORMAL)
            self.parse_btn.config(state=tk.NORMAL)

    # ═══════════════════════════════════════════════════════════
    # 💱 МЕТОДЫ ДЛЯ РАБОТЫ С КУРСОМ ВАЛЮТЫ
    # ═══════════════════════════════════════════════════════════

    def _load_settings(self):
        """Загружает настройки из JSON файла."""
        try:
            if self.settings_file.exists():
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def _save_settings(self):
        """Сохраняет текущие настройки в JSON файл."""
        try:
            # Считываем актуальные бренды из полей ввода
            brands = self.get_brands_from_ui()

            # Считываем надбавку
            try:
                markup = float(self.markup_entry.get())
            except ValueError:
                markup = 0.5

            settings = {
                "categories": self.categories_data,
                "brands": brands,
                "markup": markup,
                "threads": self.threads_var.get(),
            }
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def refresh_categories_table(self):
        """Перерисовывает таблицу категорий."""
        # Очищаем старую таблицу
        for widget in self.categories_table_frame.winfo_children():
            widget.destroy()

        # Заголовки
        tk.Label(self.categories_table_frame, text="Категория", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        tk.Label(self.categories_table_frame, text="Доставка (€)", font=("Segoe UI", 10, "bold")).grid(row=0, column=1, padx=5, pady=5)
        tk.Label(self.categories_table_frame, text="", font=("Segoe UI", 10, "bold")).grid(row=0, column=2, padx=5, pady=5)

        # Создаем записи для каждой категории
        self.category_entries = {}

        for idx, cat_data in enumerate(self.categories_data, start=1):
            # Название категории (редактируемое)
            name_entry = tk.Entry(self.categories_table_frame, font=("Segoe UI", 10), width=20)
            name_entry.insert(0, cat_data["name"])
            name_entry.grid(row=idx, column=0, padx=5, pady=2, sticky=tk.W)

            # Стоимость доставки (редактируемое)
            delivery_entry = tk.Entry(self.categories_table_frame, font=("Segoe UI", 10), width=10)
            delivery_entry.insert(0, str(cat_data["delivery"]))
            delivery_entry.grid(row=idx, column=1, padx=5, pady=2)

            # Сохраняем ссылки на entry
            self.category_entries[idx - 1] = {
                "name": name_entry,
                "delivery": delivery_entry
            }

            # Кнопка удаления
            delete_btn = tk.Button(
                self.categories_table_frame,
                text="🗑️",
                command=lambda i=idx-1: self.delete_category(i),
                bg="#f44336",
                fg="white",
                font=("Segoe UI", 9),
                width=3,
                cursor="hand2"
            )
            delete_btn.grid(row=idx, column=2, padx=5, pady=2)

    def add_category_dialog(self):
        """Диалог для добавления новой категории."""
        from tkinter import simpledialog

        # Запрашиваем название категории
        category_name = simpledialog.askstring(
            "Добавить категорию",
            "Введите название категории:",
            parent=self.root
        )

        if not category_name or not category_name.strip():
            return

        # Запрашиваем стоимость доставки
        delivery_cost = simpledialog.askstring(
            "Стоимость доставки",
            f"Введите стоимость доставки для '{category_name}' (€):",
            parent=self.root
        )

        if not delivery_cost:
            return

        try:
            delivery_cost = float(delivery_cost)
        except ValueError:
            messagebox.showerror(
                "Ошибка",
                "Неверный формат стоимости доставки!\nИспользуйте число (например: 17)"
            )
            return

        # Добавляем новую категорию
        self.categories_data.append({
            "name": category_name.strip(),
            "delivery": delivery_cost
        })

        # Обновляем таблицу
        self.refresh_categories_table()
        self._save_settings()

        messagebox.showinfo(
            "Готово!",
            f"✅ Категория '{category_name}' добавлена!"
        )

    def delete_category(self, index):
        """Удаляет категорию по индексу."""
        if len(self.categories_data) <= 1:
            messagebox.showwarning(
                "Нельзя удалить",
                "Должна остаться хотя бы одна категория!"
            )
            return

        category_name = self.categories_data[index]["name"]

        result = messagebox.askyesno(
            "Удалить категорию?",
            f"Вы уверены, что хотите удалить категорию '{category_name}'?\n\nЭто действие нельзя отменить."
        )

        if result:
            self.categories_data.pop(index)
            self.refresh_categories_table()
            self._save_settings()

            messagebox.showinfo(
                "Удалено!",
                f"✅ Категория '{category_name}' удалена!"
            )

    def refresh_brands_table(self):
        """Перерисовывает таблицу брендов."""
        for widget in self.brands_table_frame.winfo_children():
            widget.destroy()

        self.brand_entries = {}

        # Размещаем бренды в 3 колонки
        for idx, brand in enumerate(self.brands_data):
            row = idx // 3
            col = (idx % 3) * 2  # 2 ячейки на бренд (Entry + кнопка)

            entry = tk.Entry(self.brands_table_frame, font=("Segoe UI", 10), width=15)
            entry.insert(0, brand)
            entry.grid(row=row, column=col, padx=3, pady=2, sticky=tk.W)
            self.brand_entries[idx] = entry

            delete_btn = tk.Button(
                self.brands_table_frame,
                text="✕",
                command=lambda i=idx: self.delete_brand(i),
                bg="#f44336",
                fg="white",
                font=("Segoe UI", 8),
                width=2,
                cursor="hand2"
            )
            delete_btn.grid(row=row, column=col + 1, padx=(0, 10), pady=2)

    def add_brand_dialog(self):
        """Диалог добавления бренда."""
        from tkinter import simpledialog
        brand_name = simpledialog.askstring("Добавить бренд", "Введите название бренда:", parent=self.root)
        if not brand_name or not brand_name.strip():
            return
        self.brands_data.append(brand_name.strip())
        self.refresh_brands_table()
        self._save_settings()

    def delete_brand(self, index):
        """Удаляет бренд по индексу."""
        self.brands_data.pop(index)
        self.refresh_brands_table()
        self._save_settings()

    def get_brands_from_ui(self):
        """Считывает актуальные бренды из полей ввода."""
        brands = []
        for idx, entry in self.brand_entries.items():
            val = entry.get().strip()
            if val:
                brands.append(val)
        self.brands_data = brands
        return brands

    def save_category_changes(self):
        """Сохраняет изменения категорий и стоимости доставки в Excel."""
        try:
            # Сначала обновляем self.categories_data из полей ввода
            for idx, entries in self.category_entries.items():
                try:
                    new_name = entries["name"].get().strip()
                    new_delivery = float(entries["delivery"].get())

                    if not new_name:
                        raise ValueError("Название категории не может быть пустым")

                    self.categories_data[idx]["name"] = new_name
                    self.categories_data[idx]["delivery"] = new_delivery

                except ValueError as e:
                    messagebox.showerror(
                        "Ошибка",
                        f"Неверные данные в строке {idx + 1}:\n{e}\n\nИспользуйте число для стоимости доставки (например: 17)"
                    )
                    return

            if not self.file_path.exists():
                messagebox.showwarning(
                    "Файл не найден",
                    f"Excel файл {self.file_path.name} не найден!\nСоздайте файл сначала."
                )
                return

            wb = load_workbook(self.file_path)

            # Проверяем наличие листа настроек
            if "⚙️ Настройки" not in wb.sheetnames:
                messagebox.showwarning(
                    "Лист не найден",
                    "Лист '⚙️ Настройки' не найден!\nПримените формулы сначала."
                )
                wb.close()
                return

            settings_ws = wb["⚙️ Настройки"]

            # Очищаем старые данные категорий (строки 9 и далее)
            for row in range(9, 100):  # Очищаем до 100 строки
                settings_ws[f'A{row}'] = None
                settings_ws[f'B{row}'] = None

            # Записываем новые данные категорий
            thin_border = Border(
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000'),
                top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000')
            )

            for idx, cat_data in enumerate(self.categories_data, start=9):
                settings_ws[f'A{idx}'] = cat_data["name"]
                settings_ws[f'B{idx}'] = cat_data["delivery"]
                settings_ws[f'A{idx}'].border = thin_border
                settings_ws[f'B{idx}'].border = thin_border

            wb.save(self.file_path)

            # Обновляем таблицу в GUI
            self.refresh_categories_table()
            self._save_settings()

            messagebox.showinfo(
                "Сохранено!",
                f"✅ Изменения сохранены!\n\n"
                f"📋 Категорий: {len(self.categories_data)}\n\n"
                f"Изменения будут применены при следующем пересчете формул."
            )

        except Exception as e:
            messagebox.showerror(
                "Ошибка",
                f"Ошибка сохранения изменений:\n{e}"
            )

    def update_currency_rate(self):
        """Обновляет рыночный курс EUR/RUB из Yahoo Finance."""
        try:
            self.currency_rate_label.config(text="Курс EUR/RUB: загрузка...")

            # Получаем курс из Yahoo Finance (биржевой курс, близкий к Google Finance)
            import yfinance as yf

            # Тикер для пары EUR/RUB
            ticker = yf.Ticker("EURRUB=X")

            # Получаем последние данные за 1 день
            data = ticker.history(period="1d")

            if not data.empty:
                # Берем последнюю цену закрытия
                self.current_eur_rub = float(data['Close'].iloc[-1])
            else:
                raise Exception("Нет данных от Yahoo Finance")

            if self.current_eur_rub > 0:
                # Обновляем интерфейс
                self.currency_rate_label.config(
                    text=f"Курс EUR/RUB: {self.current_eur_rub:.4f} ₽ (рыночный)",
                    fg="#2196F3"
                )

                from datetime import datetime
                self.last_update_label.config(
                    text=f"Последнее обновление: {datetime.now().strftime('%d.%m.%Y %H:%M')} | Yahoo Finance"
                )

                messagebox.showinfo(
                    "Курс обновлен",
                    f"✅ Рыночный курс EUR/RUB обновлен!\n\n"
                    f"💱 {self.current_eur_rub:.4f} ₽\n"
                    f"🕐 {datetime.now().strftime('%d.%m.%Y %H:%M')}\n\n"
                    f"📊 Источник: Yahoo Finance (биржевой курс)"
                )
            else:
                raise Exception("Не удалось получить курс EUR")

        except Exception as e:
            self.currency_rate_label.config(
                text=f"Ошибка загрузки курса: {str(e)[:50]}",
                fg="#f44336"
            )
            messagebox.showerror(
                "Ошибка",
                f"Не удалось получить курс EUR/RUB:\n{e}\n\nПроверьте подключение к интернету."
            )

    def apply_currency_to_prices(self):
        """Применяет курс валюты к ценам в Excel."""
        if self.current_eur_rub <= 0:
            messagebox.showwarning(
                "Курс не загружен",
                "Сначала обновите курс валюты!\n\nНажмите '🔄 Обновить курс'"
            )
            return

        if not self.file_path.exists():
            messagebox.showwarning(
                "Файл не найден",
                f"Excel файл {self.file_path.name} не найден!"
            )
            return

        try:
            # Получаем надбавку
            markup = float(self.markup_entry.get())
            final_rate = self.current_eur_rub + markup
            use_peti = self.use_peti_coef.get()

            # Формируем текст подтверждения
            formula_text = f"Цена₽ = Цена€ × {final_rate:.2f}"
            if use_peti:
                formula_text += " × (1 + Кэф_Пети)"

            # Подтверждение
            result = messagebox.askyesno(
                "Применить курс?",
                f"Применить курс к ценам в Excel?\n\n"
                f"💱 Курс: {self.current_eur_rub:.2f} ₽\n"
                f"➕ Надбавка: {markup} ₽\n"
                f"📊 Кэф Пети: {'Включен' if use_peti else 'Выключен'}\n"
                f"═══════════════\n"
                f"📐 Формула: {formula_text}\n\n"
                + ("Кэф Пети: <15К→10%, ≤30К→9%, >30К→8%\n\n" if use_peti else "")
                + f"Все цены будут пересчитаны."
            )

            if not result:
                return

            # Загружаем Excel
            wb = load_workbook(self.file_path)
            ws = wb.active

            updated_count = 0
            total_peti_markup = 0  # Для статистики

            # Обходим строки (начиная со 2-й)
            for row_num in range(2, ws.max_row + 1):
                price_eur = ws.cell(row_num, 3).value  # C: Цена в €

                if price_eur and isinstance(price_eur, (int, float)) and price_eur > 0:
                    # Базовая цена в рублях
                    price_rub_base = price_eur * final_rate

                    # Применяем Кэф Пети, если включен
                    if use_peti:
                        # Рассчитываем коэффициент в зависимости от цены
                        if price_rub_base < 15000:
                            peti_coef = 0.10  # 10%
                        elif price_rub_base <= 30000:
                            peti_coef = 0.09  # 9%
                        else:
                            peti_coef = 0.08  # 8%

                        # Итоговая цена с наценкой
                        price_rub_final = price_rub_base * (1 + peti_coef)
                        total_peti_markup += (price_rub_final - price_rub_base)
                    else:
                        price_rub_final = price_rub_base

                    # Записываем итоговую цену
                    ws.cell(row_num, 3).value = round(price_rub_final, 2)
                    updated_count += 1

            # Сохраняем
            wb.save(self.file_path)

            # Формируем сообщение об успехе
            success_message = (
                f"✅ Цены обновлены!\n\n"
                f"📊 Обновлено товаров: {updated_count}\n"
                f"💱 Курс: {final_rate:.2f} ₽\n"
            )

            if use_peti and updated_count > 0:
                avg_peti_markup = total_peti_markup / updated_count
                success_message += (
                    f"📈 Кэф Пети: Включен\n"
                    f"💰 Средняя наценка: {avg_peti_markup:.2f} ₽\n"
                )

            success_message += "\nExcel файл сохранен."

            messagebox.showinfo("Готово!", success_message)

        except ValueError:
            messagebox.showerror(
                "Ошибка",
                "Неверный формат надбавки!\n\nВведите число (например: 0.5)"
            )
        except Exception as e:
            messagebox.showerror(
                "Ошибка",
                f"Ошибка применения курса:\n{e}"
            )

    def apply_formulas_silently(self):
        """Применяет формулы тихо (без диалоговых окон), для автоматического применения после парсинга."""
        if self.current_eur_rub <= 0:
            raise Exception("Курс EUR/RUB не загружен")

        if not self.file_path.exists():
            raise Exception(f"Excel файл {self.file_path.name} не найден")

        # Получаем текущие настройки
        markup = float(self.markup_entry.get())
        final_rate = self.current_eur_rub + markup

        # Загружаем Excel
        wb = load_workbook(self.file_path)
        ws = wb.active

        # Обновляем лист настроек с актуальным курсом
        if "⚙️ Настройки" in wb.sheetnames:
            settings_ws = wb["⚙️ Настройки"]
            settings_ws['B3'] = self.current_eur_rub
            settings_ws['B4'] = markup
        else:
            # Если листа нет, создаем его
            self._create_settings_sheet(wb, self.current_eur_rub, markup)

        # Добавляем заголовки новых столбцов (если их еще нет)
        new_headers = [
            "Доставка (₽)",      # L
            "Закупка (₽)",       # M
            "Кэф Пети (%)",      # N
            "Наш Кэф (%)",       # O
            "Цена с дост. (₽)", # P
            "Цена без дост. (₽)", # Q
            "Наша Маржа (₽)",    # R
            "Маржа Пети (₽)"     # S
        ]

        # Стили для оформления
        orange_header_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=12, name="Calibri")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        green_value_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        value_font = Font(size=11, name="Calibri")
        value_alignment = Alignment(horizontal="center", vertical="center")

        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )

        for col_idx, header in enumerate(new_headers, start=12):  # Начинаем с L (12)
            if not ws.cell(1, col_idx).value:
                ws.cell(1, col_idx).value = header

            # Применяем оранжевый стиль к заголовку
            header_cell = ws.cell(1, col_idx)
            header_cell.fill = orange_header_fill
            header_cell.font = header_font
            header_cell.alignment = header_alignment
            header_cell.border = thin_border

            # Устанавливаем ширину столбцов
            col_letter = header_cell.column_letter
            ws.column_dimensions[col_letter].width = 18

        # Заголовок "Бренд" в столбце T (20)
        if not ws.cell(1, 20).value:
            ws.cell(1, 20).value = "Бренд"
        brand_cell = ws.cell(1, 20)
        brand_cell.fill = PatternFill(start_color="2D3748", end_color="2D3748", fill_type="solid")
        brand_cell.font = Font(bold=True, color="FFFFFF", size=12, name="Calibri")
        brand_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        brand_cell.border = thin_border
        ws.column_dimensions['T'].width = 18

        processed_count = 0

        # Обходим строки с товарами (начиная со 2-й)
        for row_num in range(2, ws.max_row + 1):
            price_eur = ws.cell(row_num, 3).value  # C: Цена (€)

            # Пропускаем строки без цены
            if not price_eur:
                continue

            # === ВСТАВЛЯЕМ ФОРМУЛЫ ВМЕСТО ЗНАЧЕНИЙ ===

            # L: Доставка - VLOOKUP по категории из настроек
            formula_delivery = f"=IFERROR(VLOOKUP(F{row_num},'⚙️ Настройки'!$A$9:$B$17,2,FALSE)*'⚙️ Настройки'!$B$5,0)"

            # M: Закупка = Доставка + (Цена_EUR * Курс)
            formula_zakupka = f"=L{row_num}+(C{row_num}*'⚙️ Настройки'!$B$5)"

            # N: Кэф Пети (10%, 9%, 8% в зависимости от закупки)
            formula_peti_coef = f"=IF(M{row_num}<15000,10%,IF(M{row_num}<=30000,9%,8%))"

            # O: Наш Кэф (17%, 15%, 14%, 13% в зависимости от закупки)
            formula_nash_coef = f"=IF(M{row_num}<10000,17%,IF(M{row_num}<=20000,15%,IF(M{row_num}<=30000,14%,13%)))"

            # P: Цена с доставкой = Закупка * (1 + Кэф_Пети + Наш_Кэф)
            formula_price_with_delivery = f"=M{row_num}*(1+N{row_num}+O{row_num})"

            # Q: Цена без доставки = Цена_с_доставкой - Доставка
            formula_price_without_delivery = f"=P{row_num}-L{row_num}"

            # R: Наша Маржа = Закупка * Наш_Кэф
            formula_margin_nash = f"=M{row_num}*O{row_num}"

            # S: Маржа Пети = Закупка * Кэф_Пети
            formula_margin_peti = f"=M{row_num}*N{row_num}"

            # Вставляем формулы и применяем зеленое оформление
            formulas = [
                (12, formula_delivery),              # L: Доставка
                (13, formula_zakupka),               # M: Закупка
                (14, formula_peti_coef),             # N: Кэф Пети
                (15, formula_nash_coef),             # O: Наш Кэф
                (16, formula_price_with_delivery),   # P: Цена с доставкой
                (17, formula_price_without_delivery),# Q: Цена без доставки
                (18, formula_margin_nash),           # R: Наша Маржа
                (19, formula_margin_peti)            # S: Маржа Пети
            ]

            for col_idx, formula in formulas:
                cell = ws.cell(row_num, col_idx)
                cell.value = formula  # Вставляем формулу
                cell.fill = green_value_fill
                cell.font = value_font
                cell.alignment = value_alignment
                cell.border = thin_border
                # Формат числа для столбцов с процентами (N, O)
                if col_idx in [14, 15]:
                    cell.number_format = '0%'
                else:
                    cell.number_format = '#,##0.00'

            processed_count += 1

        # Сохраняем
        wb.save(self.file_path)
        return processed_count

    def _create_settings_sheet(self, wb, eur_rate, markup):
        """Создает лист настроек с курсом и таблицей доставки."""
        settings_ws = wb.create_sheet("⚙️ Настройки")

        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )

        # Заголовок настроек
        settings_ws['A1'] = "⚙️ НАСТРОЙКИ РАСЧЕТОВ"
        settings_ws['A1'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        settings_ws['A1'].font = Font(bold=True, color="FFFFFF", size=16, name="Calibri")
        settings_ws.merge_cells('A1:C1')

        # Курс валюты
        settings_ws['A3'] = "Курс EUR/RUB:"
        settings_ws['B3'] = eur_rate
        settings_ws['A3'].font = Font(bold=True, size=12)
        settings_ws['B3'].font = Font(size=12)
        settings_ws['B3'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        settings_ws['A4'] = "Надбавка:"
        settings_ws['B4'] = markup
        settings_ws['A4'].font = Font(bold=True, size=12)
        settings_ws['B4'].font = Font(size=12)
        settings_ws['B4'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        settings_ws['A5'] = "Итоговый курс:"
        settings_ws['B5'] = "=B3+B4"
        settings_ws['A5'].font = Font(bold=True, size=12)
        settings_ws['B5'].font = Font(bold=True, size=14)
        settings_ws['B5'].fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Таблица доставки
        settings_ws['A7'] = "📦 СТОИМОСТЬ ДОСТАВКИ (€)"
        settings_ws['A7'].font = Font(bold=True, color="FFFFFF", size=14)
        settings_ws['A7'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        settings_ws.merge_cells('A7:B7')

        settings_ws['A8'] = "Категория"
        settings_ws['B8'] = "Доставка (€)"
        settings_ws['A8'].font = Font(bold=True, size=11)
        settings_ws['B8'].font = Font(bold=True, size=11)
        settings_ws['A8'].fill = PatternFill(start_color="D0D0D0", end_color="D0D0D0", fill_type="solid")
        settings_ws['B8'].fill = PatternFill(start_color="D0D0D0", end_color="D0D0D0", fill_type="solid")

        # Таблица категорий и доставки
        delivery_table = [
            ("Очки", 12),
            ("Ракетка", 17),
            ("Кроссовки", 28),
            ("Куртка", 17),
            ("Штаны", 17),
            ("Шлем", 28),
            ("Ботинки борд", 25),
            ("Термо", 17),
            ("Очки для снега", 17)
        ]

        for idx, (cat, delivery) in enumerate(delivery_table, start=9):
            settings_ws[f'A{idx}'] = cat
            settings_ws[f'B{idx}'] = delivery
            settings_ws[f'A{idx}'].border = thin_border
            settings_ws[f'B{idx}'].border = thin_border

        # Секция БРЕНДЫ (столбец D)
        settings_ws['D1'] = "🏷️ БРЕНДЫ"
        settings_ws['D1'].font = Font(bold=True, color="FFFFFF", size=14, name="Calibri")
        settings_ws['D1'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

        settings_ws['D2'] = "Бренд"
        settings_ws['D2'].font = Font(bold=True, size=11)
        settings_ws['D2'].fill = PatternFill(start_color="D0D0D0", end_color="D0D0D0", fill_type="solid")

        for idx, brand in enumerate(self.brands_data, start=3):
            settings_ws[f'D{idx}'] = brand
            settings_ws[f'D{idx}'].border = thin_border

        # Ширина столбцов настроек
        settings_ws.column_dimensions['A'].width = 25
        settings_ws.column_dimensions['B'].width = 20
        settings_ws.column_dimensions['C'].width = 15
        settings_ws.column_dimensions['D'].width = 20

    def apply_formulas_to_excel(self):
        """Применяет все формулы расчета к товарам в Excel."""
        if self.current_eur_rub <= 0:
            messagebox.showwarning(
                "Курс не загружен",
                "Сначала обновите курс валюты!\n\nНажмите '🔄 Обновить курс'"
            )
            return

        if not self.file_path.exists():
            messagebox.showwarning(
                "Файл не найден",
                f"Excel файл {self.file_path.name} не найден!"
            )
            return

        try:
            # Получаем текущие настройки
            markup = float(self.markup_entry.get())
            final_rate = self.current_eur_rub + markup

            # Подтверждение
            result = messagebox.askyesno(
                "Применить формулы?",
                f"Применить все формулы расчета к Excel?\n\n"
                f"💱 Курс: {final_rate:.2f} ₽\n\n"
                f"Будут добавлены столбцы:\n"
                f"• L: Доставка (₽)\n"
                f"• M: Закупка (₽)\n"
                f"• N: Кэф Пети (%)\n"
                f"• O: Наш Кэф (%)\n"
                f"• P: Цена с доставкой (₽)\n"
                f"• Q: Цена без доставки (₽)\n"
                f"• R: Наша Маржа (₽)\n"
                f"• S: Маржа Пети (₽)"
            )

            if not result:
                return

            # Загружаем Excel
            wb = load_workbook(self.file_path)
            ws = wb.active

            # Обновляем лист настроек с актуальным курсом
            if "⚙️ Настройки" in wb.sheetnames:
                settings_ws = wb["⚙️ Настройки"]
                settings_ws['B3'] = self.current_eur_rub
                settings_ws['B4'] = markup
            else:
                # Если листа нет, создаем его
                self._create_settings_sheet(wb, self.current_eur_rub, markup)

            # Добавляем заголовки новых столбцов (если их еще нет)
            new_headers = [
                "Доставка (₽)",      # L
                "Закупка (₽)",       # M
                "Кэф Пети (%)",      # N
                "Наш Кэф (%)",       # O
                "Цена с дост. (₽)", # P
                "Цена без дост. (₽)", # Q
                "Наша Маржа (₽)",    # R
                "Маржа Пети (₽)"     # S
            ]

            # Стили для оформления
            orange_header_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=12, name="Calibri")
            header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            green_value_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
            value_font = Font(size=11, name="Calibri")
            value_alignment = Alignment(horizontal="center", vertical="center")

            thin_border = Border(
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000'),
                top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000')
            )

            for col_idx, header in enumerate(new_headers, start=12):  # Начинаем с L (12)
                if not ws.cell(1, col_idx).value:
                    ws.cell(1, col_idx).value = header

                # Применяем оранжевый стиль к заголовку
                header_cell = ws.cell(1, col_idx)
                header_cell.fill = orange_header_fill
                header_cell.font = header_font
                header_cell.alignment = header_alignment
                header_cell.border = thin_border

                # Устанавливаем ширину столбцов
                col_letter = header_cell.column_letter
                ws.column_dimensions[col_letter].width = 18

            # Заголовок "Бренд" в столбце T (20)
            if not ws.cell(1, 20).value:
                ws.cell(1, 20).value = "Бренд"
            brand_cell = ws.cell(1, 20)
            brand_cell.fill = PatternFill(start_color="2D3748", end_color="2D3748", fill_type="solid")
            brand_cell.font = Font(bold=True, color="FFFFFF", size=12, name="Calibri")
            brand_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            brand_cell.border = thin_border
            ws.column_dimensions['T'].width = 18

            processed_count = 0
            skipped_count = 0

            # Обходим строки с товарами (начиная со 2-й)
            for row_num in range(2, ws.max_row + 1):
                price_eur = ws.cell(row_num, 3).value  # C: Цена (€)

                # Пропускаем строки без цены
                if not price_eur:
                    skipped_count += 1
                    continue

                # === ВСТАВЛЯЕМ ФОРМУЛЫ ВМЕСТО ЗНАЧЕНИЙ ===

                # L: Доставка - VLOOKUP по категории из настроек
                formula_delivery = f"=IFERROR(VLOOKUP(F{row_num},'⚙️ Настройки'!$A$9:$B$17,2,FALSE)*'⚙️ Настройки'!$B$5,0)"

                # M: Закупка = Доставка + (Цена_EUR * Курс)
                formula_zakupka = f"=L{row_num}+(C{row_num}*'⚙️ Настройки'!$B$5)"

                # N: Кэф Пети (10%, 9%, 8% в зависимости от закупки)
                formula_peti_coef = f"=IF(M{row_num}<15000,10%,IF(M{row_num}<=30000,9%,8%))"

                # O: Наш Кэф (17%, 15%, 14%, 13% в зависимости от закупки)
                formula_nash_coef = f"=IF(M{row_num}<10000,17%,IF(M{row_num}<=20000,15%,IF(M{row_num}<=30000,14%,13%)))"

                # P: Цена с доставкой = Закупка * (1 + Кэф_Пети + Наш_Кэф)
                formula_price_with_delivery = f"=M{row_num}*(1+N{row_num}+O{row_num})"

                # Q: Цена без доставки = Цена_с_доставкой - Доставка
                formula_price_without_delivery = f"=P{row_num}-L{row_num}"

                # R: Наша Маржа = Закупка * Наш_Кэф
                formula_margin_nash = f"=M{row_num}*O{row_num}"

                # S: Маржа Пети = Закупка * Кэф_Пети
                formula_margin_peti = f"=M{row_num}*N{row_num}"

                # Вставляем формулы и применяем зеленое оформление
                formulas = [
                    (12, formula_delivery),              # L: Доставка
                    (13, formula_zakupka),               # M: Закупка
                    (14, formula_peti_coef),             # N: Кэф Пети
                    (15, formula_nash_coef),             # O: Наш Кэф
                    (16, formula_price_with_delivery),   # P: Цена с доставкой
                    (17, formula_price_without_delivery),# Q: Цена без доставки
                    (18, formula_margin_nash),           # R: Наша Маржа
                    (19, formula_margin_peti)            # S: Маржа Пети
                ]

                for col_idx, formula in formulas:
                    cell = ws.cell(row_num, col_idx)
                    cell.value = formula  # Вставляем формулу
                    cell.fill = green_value_fill
                    cell.font = value_font
                    cell.alignment = value_alignment
                    cell.border = thin_border
                    # Формат числа для столбцов с процентами (N, O)
                    if col_idx in [14, 15]:
                        cell.number_format = '0%'
                    else:
                        cell.number_format = '#,##0.00'

                processed_count += 1

            # Сохраняем
            wb.save(self.file_path)

            messagebox.showinfo(
                "Готово!",
                f"✅ Формулы применены!\n\n"
                f"📊 Обработано товаров: {processed_count}\n"
                f"⏭️ Пропущено (нет данных): {skipped_count}\n\n"
                f"Добавлены столбцы с расчетами (L-S)\n"
                f"Excel файл сохранен."
            )

        except ValueError:
            messagebox.showerror(
                "Ошибка",
                "Неверный формат надбавки!\n\nВведите число (например: 0.5)"
            )
        except Exception as e:
            messagebox.showerror(
                "Ошибка",
                f"Ошибка применения формул:\n{e}"
            )


# ═══════════════════════════════════════════════════════════
# 🚀 MAIN
# ═══════════════════════════════════════════════════════════

if __name__ == "__main__":
    root = tk.Tk()
    app = ParserApp(root)
    root.mainloop()
