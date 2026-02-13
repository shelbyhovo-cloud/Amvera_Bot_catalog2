"""
Telegram Mini App - Магазин/Каталог для группы
Всё в одном файле: бот + веб-сервер + HTML интерфейс
"""

# ═══════════════════════════════════════════════════════════
# 📦 АВТОУСТАНОВКА ЗАВИСИМОСТЕЙ
# ═══════════════════════════════════════════════════════════

import subprocess
import sys
import platform
import time
import io
from pathlib import Path

# Фикс кодировки для Windows
if platform.system() == 'Windows':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

def install_dependencies():
    """Автоматически устанавливает необходимые пакеты."""
    required_packages = {
        'aiogram': 'aiogram==3.13.1',
        'aiohttp': 'aiohttp==3.10.5',
        'openpyxl': 'openpyxl==3.1.2',
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
            print("   Попробуйте установить вручную:")
            print(f"   pip install {' '.join(missing_packages)}")
            sys.exit(1)

# Проверяем и устанавливаем зависимости при импорте
install_dependencies()


# ═══════════════════════════════════════════════════════════
# 🌐 АВТОМАТИЧЕСКИЙ ЗАПУСК SERVEO (ТУННЕЛИРОВАНИЕ)
# ═══════════════════════════════════════════════════════════

def start_serveo(port):
    """
    Запускает Serveo туннель как альтернативу ngrok.
    Возвращает (public_url, process) или (None, None) при ошибке.
    """
    print("🌐 Запускаю Serveo туннель...")
    print(f"   Порт: {port}")

    try:
        import re
        from threading import Thread

        print("   Подключаюсь к serveo.net через SSH...")

        # Запускаем SSH туннель с таймаутом
        serveo_process = subprocess.Popen(
            ['ssh', '-o', 'StrictHostKeyChecking=no',
             '-o', 'ConnectTimeout=10',
             '-o', 'ServerAliveInterval=30',
             '-o', 'ServerAliveCountMax=3',
             '-R', f'80:localhost:{port}', 'serveo.net'],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.PIPE,
            text=True,
            bufsize=1,
            creationflags=subprocess.CREATE_NO_WINDOW if platform.system() == 'Windows' else 0
        )

        serveo_url = None
        print("   Жду ответ от Serveo (макс 15 сек)...")

        # Читаем вывод и ищем URL (макс 15 секунд)
        for i in range(30):  # 30 * 0.5 = 15 секунд
            # Проверяем что процесс жив
            if serveo_process.poll() is not None:
                print("   ⚠️  Serveo процесс завершился")
                # Читаем ошибку
                output = serveo_process.stdout.read() if serveo_process.stdout else ""
                if output:
                    print(f"   Вывод: {output[:200]}")
                break

            line = serveo_process.stdout.readline()
            if line:
                # Показываем отладочную информацию
                if i < 5:  # Первые несколько строк
                    print(f"   Debug: {line.strip()[:80]}")

                # Ищем URL в формате https://xxxxx.serveo.net
                match = re.search(r'https://[a-zA-Z0-9\-]+\.serveo\.net', line)
                if match:
                    serveo_url = match.group(0)
                    break
            else:
                time.sleep(0.5)

        if serveo_url:
            print(f"✅ Serveo туннель активен!")
            print(f"🌍 Публичный URL: {serveo_url}\n")

            # Продолжаем читать вывод в фоне
            def read_output():
                try:
                    for line in serveo_process.stdout:
                        pass
                except:
                    pass

            Thread(target=read_output, daemon=True).start()

            return serveo_url, serveo_process
        else:
            print("❌ Не удалось получить URL от Serveo (таймаут 15 сек)")
            print("   Возможно Serveo перегружен или недоступен")
            try:
                serveo_process.kill()
            except:
                pass
            return None, None

    except FileNotFoundError:
        print("❌ SSH клиент не найден!")
        print("   Windows 10/11: Параметры → Приложения → Дополнительные компоненты → OpenSSH Client")
        return None, None
    except Exception as e:
        print(f"❌ Ошибка Serveo: {e}")
        import traceback
        traceback.print_exc()
        return None, None


import asyncio
import json
import logging
from aiohttp import web

from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, WebAppInfo


# ═══════════════════════════════════════════════════════════
# 🔧 УТИЛИТЫ
# ═══════════════════════════════════════════════════════════

def kill_process_on_port(port):
    """Убивает все процессы которые используют указанный порт"""
    try:
        # Находим процессы на порту
        result = subprocess.run(
            f'netstat -ano | findstr :{port}',
            shell=True,
            capture_output=True,
            text=True
        )

        if not result.stdout.strip():
            return True  # Порт свободен

        # Извлекаем PID процессов
        pids = set()
        for line in result.stdout.strip().split('\n'):
            parts = line.split()
            if len(parts) >= 5:
                pid = parts[-1]
                if pid.isdigit():
                    pids.add(pid)

        if not pids:
            return True  # Порт свободен

        # Убиваем каждый процесс
        for pid in pids:
            try:
                subprocess.run(
                    f'taskkill /PID {pid} /F',
                    shell=True,
                    capture_output=True,
                    check=True
                )
                print(f"❌ Остановил процесс на порту {port} (PID {pid})")
            except subprocess.CalledProcessError:
                pass  # Процесс уже завершён

        return True

    except Exception:
        return False


# ═══════════════════════════════════════════════════════════
# ⚙️  НАСТРОЙКИ
# ═══════════════════════════════════════════════════════════

BOT_TOKEN = "8529662300:AAHnb8e8Qh93INgnC_x3rkDc1QC20c3ulFM"
WEBAPP_HOST = "0.0.0.0"
WEBAPP_PORT = 8080

# Режим работы:
# - "auto" = автоматический туннель через Serveo (бесплатно, без регистрации)
# - "manual" = ручной режим, нужно указать свой URL ниже
MODE = "manual"

# Если MODE = "manual", вставь сюда свой HTTPS URL от Serveo/LocalTunnel/etc
MANUAL_WEBAPP_URL = "https://amvera-andrew-gurin94-run-test.amvera.io"

# WEBAPP_URL будет установлен автоматически
WEBAPP_URL = None

# ═══════════════════════════════════════════════════════════
# 📦 КАТАЛОГ ТОВАРОВ (можно редактировать)
# ═══════════════════════════════════════════════════════════

# Стандартные товары (используются если нет Excel файла)
PRODUCTS_DEFAULT = [
    {
        "id": 1,
        "name": "Футболка Premium",
        "description": "100% хлопок, удобная посадка",
        "price": 1500,
        "image": "👕",
    },
    {
        "id": 2,
        "name": "Кроссовки Sport",
        "description": "Лёгкие беговые кроссовки",
        "price": 4500,
        "image": "👟",
    },
    {
        "id": 3,
        "name": "Рюкзак Urban",
        "description": "Городской рюкзак 20L с USB",
        "price": 2800,
        "image": "🎒",
    },
    {
        "id": 4,
        "name": "Наушники Pro",
        "description": "Беспроводные с шумоподавлением",
        "price": 6000,
        "image": "🎧",
    },
    {
        "id": 5,
        "name": "Смарт-часы",
        "description": "Фитнес-трекер + уведомления",
        "price": 8500,
        "image": "⌚",
    },
    {
        "id": 6,
        "name": "Кепка Classic",
        "description": "Бейсболка с логотипом",
        "price": 900,
        "image": "🧢",
    },
]

PRODUCTS = []  # Будет загружено из Excel или использованы стандартные

def load_products_from_excel(file_path=None):
    """Загружает товары из Excel файла."""
    global PRODUCTS

    # Если путь не указан, ищем в папке со скриптом
    if file_path is None:
        script_dir = Path(__file__).parent
        file_path = script_dir / "products_links.xlsx"
    else:
        file_path = Path(file_path)

    if not file_path.exists():
        print(f"📦 Excel файл не найден: {file_path}")
        print("   Используются стандартные товары")
        print("   Для управления товарами через Excel:")
        print("   1. Запусти: python parser_gui.py")
        print("   2. Создай шаблон и заполни ссылки")
        print("   3. Спарси товары")
        print("   4. Перезапусти мини-апп\n")
        PRODUCTS = PRODUCTS_DEFAULT
        return

    try:
        # Проверяем openpyxl
        try:
            import openpyxl
        except ImportError:
            print("📦 Устанавливаю openpyxl...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl'])
            import openpyxl

        from openpyxl import load_workbook

        wb = load_workbook(file_path)
        ws = wb.active

        products = []

        for row_num in range(2, ws.max_row + 1):
            name = ws.cell(row_num, 2).value          # B: Название
            price = ws.cell(row_num, 3).value         # C: Цена
            description = ws.cell(row_num, 4).value   # D: Описание
            category = ws.cell(row_num, 5).value      # E: Группа
            subcategory = ws.cell(row_num, 6).value   # F: Подгруппа
            local_images = ws.cell(row_num, 8).value  # H: Локальное фото

            # Пропускаем строки без данных
            if not name or not price:
                continue

            # Определяем изображение для показа
            image_to_use = "📦"  # По умолчанию placeholder эмодзи

            # Если есть локальные фотографии, используем первую
            if local_images:
                # Локальные фото могут быть разделены запятыми
                local_photos = [img.strip() for img in local_images.split(',')]
                if local_photos:
                    # Убираем префикс "images\" или "images/" если он есть
                    photo_path = local_photos[0].replace('images\\', '').replace('images/', '')
                    # Используем первую фотографию
                    image_to_use = f"/images/{photo_path}"

            products.append({
                "id": row_num - 1,
                "name": name,
                "description": description or "",
                "price": int(price) if price else 0,
                "image": image_to_use,
                "category": category or "",
                "subcategory": subcategory or "",
            })

        if products:
            PRODUCTS = products
            print(f"✅ Загружено товаров из Excel: {len(products)}")

            # Подсчитываем товары с фотографиями
            with_photos = sum(1 for p in products if p['image'].startswith('/images/'))
            print(f"   📸 Товаров с фотографиями: {with_photos}")
            print(f"   📦 Товаров с эмодзи: {len(products) - with_photos}\n")
        else:
            print("⚠️  Excel файл пустой, используются стандартные товары\n")
            PRODUCTS = PRODUCTS_DEFAULT

    except Exception as e:
        print(f"❌ Ошибка загрузки Excel: {e}")
        print("   Используются стандартные товары\n")
        PRODUCTS = PRODUCTS_DEFAULT

# ═══════════════════════════════════════════════════════════
# 🤖 TELEGRAM БОТ
# ═══════════════════════════════════════════════════════════

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()


@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    """Команда /start - показывает приветствие и кнопку магазина."""
    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="🛍 Открыть магазин",
                    web_app=WebAppInfo(url=WEBAPP_URL),
                )
            ]
        ]
    )
    await message.answer(
        "🎉 <b>Добро пожаловать в наш магазин!</b>\n\n"
        "Нажми кнопку ниже, чтобы открыть каталог товаров и сделать заказ.",
        reply_markup=keyboard,
        parse_mode="HTML",
    )


@dp.message(Command("shop"))
async def cmd_shop(message: types.Message):
    """Команда /shop - открывает магазин."""
    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="🛍 Открыть магазин",
                    web_app=WebAppInfo(url=WEBAPP_URL),
                )
            ]
        ]
    )
    await message.answer(
        "Нажми кнопку, чтобы открыть каталог:",
        reply_markup=keyboard,
    )


@dp.message(Command("reload"))
async def cmd_reload(message: types.Message):
    """Команда /reload - перезагружает каталог товаров из Excel."""
    await message.answer("🔄 Перезагружаю каталог товаров...")

    try:
        load_products_from_excel()
        await message.answer(
            f"✅ Каталог обновлён!\n\n"
            f"📦 Товаров: {len(PRODUCTS)}\n"
            f"📸 С фото: {sum(1 for p in PRODUCTS if p['image'].startswith('/images/'))}"
        )
    except Exception as e:
        await message.answer(f"❌ Ошибка обновления каталога:\n{str(e)}")


@dp.message(F.document)
async def handle_document(message: types.Message):
    """Обрабатывает загрузку архивов с каталогом товаров."""
    document = message.document

    # Проверяем расширение файла (только ZIP для простоты)
    if not document.file_name.endswith('.zip'):
        await message.answer(
            "⚠️ Пожалуйста, отправь ZIP архив с каталогом.\n\n"
            "📝 Как создать архив:\n"
            "  1. Положи в одну папку:\n"
            "     • products_links.xlsx\n"
            "     • папку images/\n"
            "  2. Выдели оба → ПКМ → Отправить → Сжатая ZIP-папка\n\n"
            "Структура архива:\n"
            "  📁 catalog.zip\n"
            "     ├── 📄 products_links.xlsx\n"
            "     └── 📁 images/\n"
            "          ├── 🖼 product_1.webp\n"
            "          ├── 🖼 product_2.webp\n"
            "          └── ..."
        )
        return

    try:
        await message.answer("📥 Скачиваю архив...")

        # Скачиваем файл
        script_dir = Path(__file__).parent
        archive_path = script_dir / document.file_name

        await bot.download(document, destination=archive_path)
        await message.answer("✅ Архив скачан, распаковываю...")

        # Распаковываем ZIP
        import zipfile
        with zipfile.ZipFile(archive_path, 'r') as zip_ref:
            zip_ref.extractall(script_dir)

        # Удаляем архив
        archive_path.unlink()

        await message.answer("✅ Архив распакован, обновляю каталог...")

        # Перезагружаем товары
        load_products_from_excel()

        await message.answer(
            f"🎉 Каталог успешно обновлён!\n\n"
            f"📦 Товаров: {len(PRODUCTS)}\n"
            f"📸 С фотографиями: {sum(1 for p in PRODUCTS if p['image'].startswith('/images/'))}\n\n"
            f"Используй /shop чтобы открыть магазин"
        )

    except zipfile.BadZipFile:
        await message.answer("❌ Ошибка: файл повреждён или это не ZIP архив")
    except Exception as e:
        logger.error("Ошибка обработки архива: %s", e)
        await message.answer(f"❌ Ошибка обработки архива:\n{str(e)}")


@dp.message(F.web_app_data)
async def handle_web_app_data(message: types.Message):
    """Обрабатывает заказ, полученный из Mini App."""
    try:
        data = json.loads(message.web_app_data.data)
        items = data.get("items", [])
        total = data.get("total", 0)

        if not items:
            await message.answer("❌ Корзина пуста!")
            return

        # Формируем красивое сообщение с заказом
        order_text = "📦 <b>Новый заказ!</b>\n\n"
        for item in items:
            subtotal = item["price"] * item["quantity"]
            order_text += (
                f"  {item.get('image', '▪️')} <b>{item['name']}</b>\n"
                f"     {item['quantity']} шт. × {item['price']} ₽ = {subtotal} ₽\n\n"
            )

        order_text += f"💰 <b>Итого: {total} ₽</b>\n"
        order_text += f"👤 Покупатель: {message.from_user.full_name}"

        if message.from_user.username:
            order_text += f" (@{message.from_user.username})"

        await message.answer(order_text, parse_mode="HTML")

        # Логируем в консоль
        logger.info(
            "Заказ от %s (@%s): %d товаров на %d ₽",
            message.from_user.full_name,
            message.from_user.username or "без username",
            len(items),
            total,
        )

        # Опционально: отправить в группу/канал
        # CHANNEL_ID = -1001234567890  # ID канала/группы
        # await bot.send_message(CHANNEL_ID, order_text, parse_mode="HTML")

    except (json.JSONDecodeError, KeyError) as e:
        logger.error("Ошибка обработки заказа: %s", e)
        await message.answer("❌ Произошла ошибка при обработке заказа.")


# ═══════════════════════════════════════════════════════════
# 🌐 ВЕБ-СЕРВЕР (раздаёт HTML и API)
# ═══════════════════════════════════════════════════════════

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Магазин</title>
    <script src="https://telegram.org/js/telegram-web-app.js"></script>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: var(--tg-theme-bg-color, #ffffff);
            color: var(--tg-theme-text-color, #000000);
            padding: 16px;
            padding-bottom: 80px;
        }

        h1 {
            font-size: 24px;
            margin-bottom: 8px;
            color: var(--tg-theme-text-color);
        }

        .subtitle {
            color: var(--tg-theme-hint-color, #999);
            margin-bottom: 20px;
            font-size: 14px;
        }

        .products-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
            gap: 12px;
            margin-bottom: 20px;
        }

        .product-card {
            background: var(--tg-theme-secondary-bg-color, #f4f4f5);
            border-radius: 12px;
            padding: 12px;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
            border: 2px solid transparent;
        }

        .product-card:active {
            transform: scale(0.97);
        }

        .product-card.in-cart {
            border-color: var(--tg-theme-button-color, #3390ec);
        }

        .product-image {
            font-size: 48px;
            text-align: center;
            margin-bottom: 8px;
            min-height: 60px;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .product-image img {
            width: 100%;
            height: 120px;
            object-fit: cover;
            border-radius: 8px;
        }

        .product-name {
            font-weight: 600;
            font-size: 14px;
            margin-bottom: 4px;
            color: var(--tg-theme-text-color);
        }

        .product-description {
            font-size: 12px;
            color: var(--tg-theme-hint-color, #999);
            margin-bottom: 8px;
            line-height: 1.3;
        }

        .product-price {
            font-size: 16px;
            font-weight: 700;
            color: var(--tg-theme-button-color, #3390ec);
        }

        .product-quantity {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-top: 8px;
            gap: 8px;
        }

        .quantity-btn {
            width: 32px;
            height: 32px;
            border-radius: 8px;
            border: none;
            background: var(--tg-theme-button-color, #3390ec);
            color: var(--tg-theme-button-text-color, #ffffff);
            font-size: 18px;
            font-weight: bold;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .quantity-btn:active {
            opacity: 0.7;
        }

        .quantity-display {
            font-weight: 600;
            font-size: 16px;
            min-width: 24px;
            text-align: center;
        }

        .cart-footer {
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            background: var(--tg-theme-secondary-bg-color, #f4f4f5);
            padding: 12px 16px;
            box-shadow: 0 -2px 10px rgba(0,0,0,0.1);
            display: none;
        }

        .cart-footer.visible {
            display: block;
        }

        .cart-summary {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 8px;
            font-size: 14px;
        }

        .cart-total {
            font-size: 20px;
            font-weight: 700;
            color: var(--tg-theme-button-color, #3390ec);
        }

        .order-btn {
            width: 100%;
            padding: 12px;
            border-radius: 10px;
            border: none;
            background: var(--tg-theme-button-color, #3390ec);
            color: var(--tg-theme-button-text-color, #ffffff);
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
        }

        .order-btn:active {
            opacity: 0.8;
        }

        .empty-cart {
            text-align: center;
            padding: 40px 20px;
            color: var(--tg-theme-hint-color, #999);
        }

        .empty-cart-icon {
            font-size: 64px;
            margin-bottom: 16px;
        }
    </style>
</head>
<body>
    <h1>🛍 Наш магазин</h1>
    <p class="subtitle">Выберите товары и добавьте в корзину</p>

    <div class="products-grid" id="productsGrid"></div>

    <div class="cart-footer" id="cartFooter">
        <div class="cart-summary">
            <span>Товаров: <span id="cartCount">0</span></span>
            <span class="cart-total"><span id="cartTotal">0</span> ₽</span>
        </div>
        <button class="order-btn" id="orderBtn">Оформить заказ</button>
    </div>

    <script>
        const tg = window.Telegram.WebApp;
        tg.expand();
        tg.MainButton.hide();

        let cart = {};
        let products = [];

        // Загружаем товары с сервера
        fetch('/api/products')
            .then(res => res.json())
            .then(data => {
                products = data;
                renderProducts();
            });

        function renderProducts() {
            const grid = document.getElementById('productsGrid');
            grid.innerHTML = '';

            products.forEach(product => {
                const card = document.createElement('div');
                card.className = 'product-card';
                if (cart[product.id]) {
                    card.classList.add('in-cart');
                }

                const quantity = cart[product.id] || 0;

                // Определяем как показывать изображение
                let imageHtml;
                if (product.image.startsWith('/images/')) {
                    // Реальная фотография
                    imageHtml = `<img src="${product.image}" alt="${product.name}" onerror="this.parentElement.innerHTML='📦'">`;
                } else {
                    // Placeholder эмодзи
                    imageHtml = product.image;
                }

                card.innerHTML = `
                    <div class="product-image">${imageHtml}</div>
                    <div class="product-name">${product.name}</div>
                    <div class="product-description">${product.description}</div>
                    <div class="product-price">${product.price} ₽</div>
                    ${quantity > 0 ? `
                        <div class="product-quantity">
                            <button class="quantity-btn" onclick="changeQuantity(${product.id}, -1)">−</button>
                            <span class="quantity-display">${quantity}</span>
                            <button class="quantity-btn" onclick="changeQuantity(${product.id}, 1)">+</button>
                        </div>
                    ` : ''}
                `;

                if (quantity === 0) {
                    card.onclick = () => changeQuantity(product.id, 1);
                }

                grid.appendChild(card);
            });

            updateCartFooter();
        }

        function changeQuantity(productId, delta) {
            if (!cart[productId]) {
                cart[productId] = 0;
            }

            cart[productId] += delta;

            if (cart[productId] <= 0) {
                delete cart[productId];
            }

            renderProducts();
            tg.HapticFeedback.impactOccurred('light');
        }

        function updateCartFooter() {
            const footer = document.getElementById('cartFooter');
            const cartCount = document.getElementById('cartCount');
            const cartTotal = document.getElementById('cartTotal');

            let totalItems = 0;
            let totalPrice = 0;

            for (const [productId, quantity] of Object.entries(cart)) {
                const product = products.find(p => p.id === parseInt(productId));
                if (product) {
                    totalItems += quantity;
                    totalPrice += product.price * quantity;
                }
            }

            if (totalItems > 0) {
                footer.classList.add('visible');
                cartCount.textContent = totalItems;
                cartTotal.textContent = totalPrice;
            } else {
                footer.classList.remove('visible');
            }
        }

        document.getElementById('orderBtn').addEventListener('click', () => {
            const items = [];
            let total = 0;

            for (const [productId, quantity] of Object.entries(cart)) {
                const product = products.find(p => p.id === parseInt(productId));
                if (product) {
                    items.push({
                        id: product.id,
                        name: product.name,
                        price: product.price,
                        quantity: quantity,
                        image: product.image
                    });
                    total += product.price * quantity;
                }
            }

            const orderData = {
                items: items,
                total: total
            };

            tg.sendData(JSON.stringify(orderData));
            tg.close();
        });
    </script>
</body>
</html>
"""


async def handle_index(request: web.Request) -> web.Response:
    """Отдаёт HTML страницу Mini App."""
    return web.Response(text=HTML_TEMPLATE, content_type="text/html")


async def handle_products(request: web.Request) -> web.Response:
    """API: список товаров в формате JSON."""
    return web.json_response(PRODUCTS)


def create_web_app() -> web.Application:
    """Создаёт веб-приложение aiohttp."""
    app = web.Application()
    app.router.add_get("/", handle_index)
    app.router.add_get("/api/products", handle_products)

    # Раздаём статические файлы (фотографии товаров)
    script_dir = Path(__file__).parent
    images_dir = script_dir / "images"
    if images_dir.exists():
        app.router.add_static("/images/", path=images_dir, name="images")
        logger.info(f"📁 Раздача изображений из: {images_dir}")

    return app


# ═══════════════════════════════════════════════════════════
# 🚀 ЗАПУСК
# ═══════════════════════════════════════════════════════════

async def main():
    """Запускает бота и веб-сервер одновременно."""
    global WEBAPP_URL

    # Освобождаем порт перед запуском
    print(f"🔍 Проверяю порт {WEBAPP_PORT}...")
    kill_process_on_port(WEBAPP_PORT)
    print(f"✅ Порт {WEBAPP_PORT} готов к использованию\n")

    # Загружаем товары из Excel
    load_products_from_excel()

    tunnel_process = None

    # 1. Настраиваем публичный URL
    if MODE == "auto":
        # Автоматический режим с Serveo
        print("🔧 Режим: Автоматический (Serveo)\n")
        WEBAPP_URL, tunnel_process = start_serveo(WEBAPP_PORT)

        if not WEBAPP_URL:
            # Serveo не сработал - запускаемся локально
            print("\n" + "=" * 60)
            print("⚠️  SERVEO НЕДОСТУПЕН - ЗАПУСК В ЛОКАЛЬНОМ РЕЖИМЕ")
            print("=" * 60)
            print()
            print("🏠 Бот запущен локально на http://localhost:8080")
            print()
            print("⚠️  ВАЖНО:")
            print("   • Telegram Mini App НЕ БУДЕТ РАБОТАТЬ")
            print("   • Можно открыть http://localhost:8080 в браузере")
            print("   • Для полной работы нужен публичный HTTPS URL")
            print()
            print("💡 Как получить публичный URL:")
            print()
            print("   ВАРИАНТ 1: Serveo (ручной режим)")
            print("     1. Открой новый терминал")
            print(f"     2. Запусти: ssh -R 80:localhost:{WEBAPP_PORT} serveo.net")
            print("     3. Скопируй полученный URL")
            print("     4. Вставь URL в mini_app.py (строка 205):")
            print('        MANUAL_WEBAPP_URL = "твой_url"')
            print("     5. Измени MODE = \"manual\" (строка 202)")
            print("     6. Перезапусти бота")
            print()
            print("   ВАРИАНТ 2: LocalTunnel")
            print(f"     npx localtunnel --port {WEBAPP_PORT}")
            print()
            print("   ВАРИАНТ 3: Деплой на облако (Railway, Render)")
            print("     Бот будет работать 24/7 с автоматическим HTTPS")
            print()
            print("=" * 60)
            print()

            # Запускаемся локально для тестирования
            WEBAPP_URL = f"http://localhost:{WEBAPP_PORT}"
            print(f"▶️  Запускаю в локальном режиме...")
            print(f"   Адрес: {WEBAPP_URL}")
            print()

            # Автоматически открываем браузер через 3 секунды
            import webbrowser
            from threading import Timer
            def open_browser():
                try:
                    webbrowser.open(WEBAPP_URL)
                    print("🌐 Открыл веб-интерфейс в браузере")
                except:
                    pass
            Timer(3.0, open_browser).start()

    else:
        # Ручной режим - используем указанный URL
        WEBAPP_URL = MANUAL_WEBAPP_URL
        print("📌 Ручной режим: используется URL из настроек")
        print(f"🌍 URL: {WEBAPP_URL}\n")

    # 2. Запускаем веб-сервер
    web_app = create_web_app()
    runner = web.AppRunner(web_app)
    await runner.setup()
    site = web.TCPSite(runner, WEBAPP_HOST, WEBAPP_PORT)
    await site.start()

    logger.info("=" * 60)
    logger.info("🌐 Локальный сервер: http://%s:%s", WEBAPP_HOST, WEBAPP_PORT)
    logger.info("🌍 Публичный URL (Mini App): %s", WEBAPP_URL)
    logger.info("=" * 60)

    # 3. Запускаем бота
    logger.info("🤖 Telegram бот запущен!")
    logger.info("💬 Напиши боту /start чтобы открыть магазин!\n")

    try:
        await dp.start_polling(bot)
    finally:
        # Останавливаем всё при выходе
        logger.info("Останавливаю сервер...")
        await runner.cleanup()
        if tunnel_process:
            logger.info("Останавливаю туннель...")
            tunnel_process.kill()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Остановка бота...")
