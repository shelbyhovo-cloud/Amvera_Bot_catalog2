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
import io
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

# ═══════════════════════════════════════════════════════════
# 📄 СОЗДАНИЕ КРАСИВОГО ШАБЛОНА EXCEL
# ═══════════════════════════════════════════════════════════

def create_beautiful_template(file_path=None):
    """Создаёт красиво оформленный шаблон Excel."""

    if file_path is None:
        script_dir = Path(__file__).parent
        file_path = script_dir / "products_links.xlsx"
    else:
        file_path = Path(file_path)

    wb = Workbook()
    ws = wb.active
    ws.title = "🛍 Товары"

    # Заголовки (без эмодзи)
    headers = ["URL товара", "Название", "Цена (€)", "Описание", "Группа", "Подгруппа", "URL фото", "Локальное фото", "Размеры", "Последнее обновление", "Статус"]
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
    ws.column_dimensions['D'].width = 45  # Описание
    ws.column_dimensions['E'].width = 18  # Группа
    ws.column_dimensions['F'].width = 18  # Подгруппа
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

        # Цвет фона строки
        row_fill = PatternFill(start_color=row_colors[idx % 2], end_color=row_colors[idx % 2], fill_type="solid")

        # Применяем стили к каждой ячейке
        for col_num, cell in enumerate(ws[row_num], start=1):
            cell.fill = row_fill
            cell.font = data_font
            cell.border = data_border

            # Выравнивание по центру для определённых колонок
            if col_num in [3, 7, 10, 12]:  # Цена, Эмодзи, Размеры, Статус
                cell.alignment = center_alignment
            else:
                cell.alignment = data_alignment

        # Высота строки
        ws.row_dimensions[row_num].height = 25

    # Закрепляем первую строку (заголовки)
    ws.freeze_panes = "A2"

    # Автофильтр (теперь до колонки K)
    ws.auto_filter.ref = f"A1:K{ws.max_row}"

    wb.save(file_path)
    return file_path


# ═══════════════════════════════════════════════════════════
# 🕷️ ПАРСИНГ (из update_products.py)
# ═══════════════════════════════════════════════════════════

def download_image(image_url, save_dir, product_id):
    """Скачивает изображение и сохраняет локально (с проверкой существования)."""
    try:
        # Генерируем имя файла заранее (используем hash URL для уникальности)
        url_hash = hash(image_url) % 10000

        # Пробуем найти существующий файл с таким хешем (любое расширение)
        for ext in ['.jpg', '.png', '.webp', '.jpeg']:
            filename = f"product_{product_id}_{url_hash}{ext}"
            filepath = save_dir / filename

            # Если файл уже существует - пропускаем скачивание
            if filepath.exists():
                return str(filepath.relative_to(save_dir.parent))

        # Если файла нет - скачиваем
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

        # Генерируем имя файла
        filename = f"product_{product_id}_{url_hash}{ext}"
        filepath = save_dir / filename

        # Сохраняем файл
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

        price_match = re.search(r'data-price="([^"]+)"', html, re.IGNORECASE)
        if price_match:
            try:
                price = float(price_match.group(1))
            except:
                price = 0
        else:
            price = 0

        desc_match = re.search(r'<meta name="description" content="([^"]+)"', html, re.IGNORECASE)
        description = desc_match.group(1)[:100] if desc_match else ""

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

        # Скачиваем фотки
        images_dir = script_dir / "images"
        images_dir.mkdir(exist_ok=True)

        local_images = []
        for img_url in image_urls:
            local_path = download_image(img_url, images_dir, product_id)
            if local_path:
                local_images.append(local_path)

        return {
            "name": name,
            "description": description,
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

        response = requests.get(url, timeout=10)
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

                    # Скачиваем фотки
                    images_dir = script_dir / "images"
                    images_dir.mkdir(exist_ok=True)

                    local_images = []
                    for img_url in image_urls:
                        local_path = download_image(img_url, images_dir, product_id)
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
        desc_match = re.search(r'<meta property="og:description" content="([^"]+)"', html)
        price_match = re.search(r'<meta property="product:price:amount" content="([^"]+)"', html)

        name = name_match.group(1) if name_match else "Без названия"
        description = desc_match.group(1)[:100] if desc_match else ""

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

        # Скачиваем фотки
        images_dir = script_dir / "images"
        images_dir.mkdir(exist_ok=True)

        local_images = []
        for img_url in image_urls:
            local_path = download_image(img_url, images_dir, product_id)
            if local_path:
                local_images.append(local_path)

        return {
            "name": name,
            "description": description,
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

        # Стили
        style = ttk.Style()
        style.theme_use('clam')

        # Главный фрейм
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок
        title_label = tk.Label(
            main_frame,
            text="🕷️ Парсер товаров для Telegram магазина",
            font=("Segoe UI", 16, "bold"),
            fg="#1F4E78"
        )
        title_label.pack(pady=(0, 20))

        # Фрейм для кнопок
        button_frame = ttk.Frame(main_frame)
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

        # Информационная панель
        info_frame = ttk.LabelFrame(main_frame, text="ℹ️ Информация", padding="10")
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
        log_frame = ttk.LabelFrame(main_frame, text="📋 Журнал работы", padding="10")
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

    def log(self, message):
        """Добавляет сообщение в лог."""
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

            file_path = create_beautiful_template(self.file_path)
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

            images_dir = self.script_dir / "images"
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
        """Парсит товары из Excel."""
        # Блокируем кнопки
        self.create_btn.config(state=tk.DISABLED)
        self.parse_btn.config(state=tk.DISABLED)

        self.update_status("🕷️ Парсинг в процессе...")
        self.log("\n" + "=" * 80)
        self.log("🕷️ ПАРСИНГ ТОВАРОВ")
        self.log("=" * 80)
        self.log("")

        try:
            wb = load_workbook(self.file_path)
            ws = wb.active

            updated_count = 0
            error_count = 0
            total_rows = ws.max_row - 1  # Минус заголовок

            for row_num in range(2, ws.max_row + 1):
                url = ws.cell(row_num, 1).value

                if not url or not url.startswith("http"):
                    self.log(f"[{row_num - 1}/{total_rows}] ⏭️ Пропущено (нет URL)")
                    ws.cell(row_num, 11).value = "Пропущено (нет URL)"  # K: Статус
                    continue

                # Обновляем статус
                self.update_status(f"🕷️ Парсинг товара {row_num - 1}/{total_rows}...")
                self.log(f"[{row_num - 1}/{total_rows}] 🔍 Парсинг: {url[:60]}...")

                product_id = row_num - 1
                product_data, error = parse_product(url, self.script_dir, product_id)

                if error:
                    self.log(f"    ❌ {error}")
                    ws.cell(row_num, 11).value = error  # K: Статус
                    error_count += 1
                else:
                    self.log(f"    ✅ {product_data['name']}")
                    self.log(f"       💰 Цена: {product_data['price']} €")

                    if product_data.get('image_urls'):
                        photos_count = len(product_data['image_urls'].split(','))
                        self.log(f"       📷 Фото: {photos_count} шт.")

                    ws.cell(row_num, 2).value = product_data['name']           # B: Название
                    ws.cell(row_num, 3).value = product_data['price']          # C: Цена
                    ws.cell(row_num, 4).value = product_data['description']    # D: Описание
                    # E: Группа (заполняется вручную)
                    # F: Подгруппа (заполняется вручную)
                    ws.cell(row_num, 7).value = product_data['image_urls']     # G: URL фото
                    ws.cell(row_num, 8).value = product_data['local_images']   # H: Локальное фото
                    ws.cell(row_num, 9).value = product_data.get('sizes', '')  # I: Размеры (автоматически!)
                    ws.cell(row_num, 10).value = datetime.now().strftime("%Y-%m-%d %H:%M")  # J: Обновление
                    ws.cell(row_num, 11).value = "✅ Обновлено"                # K: Статус

                    updated_count += 1

                # Задержка
                import time
                time.sleep(2)

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

            self.update_status(f"✅ Парсинг завершён: {updated_count} товаров обновлено")

            messagebox.showinfo(
                "Парсинг завершён",
                f"✅ Обновлено товаров: {updated_count}\n❌ Ошибок: {error_count}\n\n📄 {self.file_path}\n📁 Фотки: {self.script_dir / 'images'}"
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
# 🚀 MAIN
# ═══════════════════════════════════════════════════════════

if __name__ == "__main__":
    root = tk.Tk()
    app = ParserApp(root)
    root.mainloop()
