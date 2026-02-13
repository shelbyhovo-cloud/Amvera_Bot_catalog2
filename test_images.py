"""
Тестовый скрипт для проверки парсинга фоток
"""

import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import requests
import re
from pathlib import Path

url = "https://www.tradeinn.com/volleyball/ru/asics-%D0%9E%D0%B1%D1%83%D0%B2%D1%8C-%D0%B4%D0%BB%D1%8F-%D0%B7%D0%B0%D0%BA%D1%80%D1%8B%D1%82%D1%8B%D1%85-%D0%BA%D0%BE%D1%80%D1%82%D0%BE%D0%B2-netburner-ballistic-ff-3/141608258/p"

print("🧪 ТЕСТИРУЮ ПАРСИНГ ФОТОК")
print("=" * 70)
print(f"URL: {url}\n")

# Скачиваем страницу
print("📥 Скачиваю страницу...")
response = requests.get(url, timeout=10)
html = response.text
print(f"✅ Страница скачана ({len(html)} байт)\n")

# ====================================================================
# МЕТОД 1: Поиск JSON с данными товара
# ====================================================================
print("🔍 МЕТОД 1: Ищу JSON с данными...")
json_match = re.search(r'var\s+product\s*=\s*(\{[^}]+images[^}]+\})', html, re.DOTALL)
if json_match:
    print(f"   Найден JSON: {json_match.group(1)[:100]}...")
else:
    print("   ❌ JSON не найден")

json_match2 = re.search(r'"images"\s*:\s*(\[[^\]]+\])', html, re.DOTALL)
if json_match2:
    print(f"   Найден images массив: {json_match2.group(1)[:100]}...")
else:
    print("   ❌ images массив не найден")

# ====================================================================
# МЕТОД 2: Поиск data-zoom-image
# ====================================================================
print("\n🔍 МЕТОД 2: Ищу data-zoom-image...")
zoom_images = re.findall(r'data-zoom-image="([^"]+)"', html, re.IGNORECASE)
print(f"   Найдено: {len(zoom_images)} шт.")
for i, img in enumerate(zoom_images[:3], 1):
    print(f"   {i}. {img[:80]}...")

# ====================================================================
# МЕТОД 3: Поиск всех URL с /f/числа/числа/
# ====================================================================
print("\n🔍 МЕТОД 3: Ищу все URL с pattern /f/\\d+/\\d+/...")
all_images = re.findall(r'https://[^"\']+/f/\d+/\d+/[^"\']+\.(?:jpg|jpeg|png|webp)', html, re.IGNORECASE)
print(f"   Найдено: {len(all_images)} шт.")

# Фильтруем дубликаты
unique_images = list(set(all_images))
print(f"   Уникальных: {len(unique_images)} шт.\n")

for i, img in enumerate(unique_images[:5], 1):
    print(f"   {i}. {img}")

# ====================================================================
# МЕТОД 3.5: Ищу srcset и data-large атрибуты
# ====================================================================
print("\n🔍 МЕТОД 3.5: Ищу srcset и data-large...")
srcset_matches = re.findall(r'srcset=["\']([^"\']+)["\']', html, re.IGNORECASE)
print(f"   srcset найдено: {len(srcset_matches)} шт.")
if srcset_matches:
    for i, src in enumerate(srcset_matches[:3], 1):
        print(f"   {i}. {src[:100]}...")

data_large = re.findall(r'data-large[^=]*=["\']([^"\']+)["\']', html, re.IGNORECASE)
print(f"   data-large найдено: {len(data_large)} шт.")
if data_large:
    for i, dl in enumerate(data_large[:3], 1):
        print(f"   {i}. {dl}")

# ====================================================================
# МЕТОД 3.6: Попробуем модифицировать URL для получения больших версий
# ====================================================================
print("\n🔍 МЕТОД 3.6: Пробую модифицировать URL...")
if unique_images:
    test_url = unique_images[0]
    print(f"   Базовый URL: {test_url}")

    # Попробуем разные модификации
    modifications = []

    # Вариант 1: заменить путь на /images/
    mod1 = test_url.replace('/f/', '/images/')
    modifications.append(("Путь /images/", mod1))

    # Вариант 2: добавить _large перед расширением
    base, ext = test_url.rsplit('.', 1)
    mod2 = f"{base}_large.{ext}"
    modifications.append(("Суффикс _large", mod2))

    # Вариант 3: добавить _xl
    mod3 = f"{base}_xl.{ext}"
    modifications.append(("Суффикс _xl", mod3))

    # Вариант 4: заменить на /800/ или /1200/
    mod4 = test_url.replace('/f/', '/800/')
    modifications.append(("Путь /800/", mod4))

    mod5 = test_url.replace('/f/', '/1200/')
    modifications.append(("Путь /1200/", mod5))

    # Вариант 6: Попробовать убрать числовое имя файла и оставить только название товара
    if 'asics' not in test_url.lower():
        # Берём URL с названием товара
        for img in unique_images:
            if 'asics' in img.lower() or '-' in img:
                test_url_with_name = img
                modifications.append(("URL с названием", test_url_with_name))
                break

    print(f"   Тестирую {len(modifications)} модификаций...")
    for name, mod_url in modifications:
        try:
            resp = requests.head(mod_url, timeout=5)
            if resp.status_code == 200:
                size = int(resp.headers.get('content-length', 0)) / 1024
                print(f"   ✅ {name}: {size:.1f} KB - {mod_url}")
            else:
                print(f"   ❌ {name}: HTTP {resp.status_code}")
        except Exception as e:
            print(f"   ❌ {name}: {str(e)[:50]}")

# ====================================================================
# МЕТОД 4: Скачиваем ВСЕ найденные фотки для проверки размеров
# ====================================================================
if unique_images:
    print(f"\n📷 Скачиваю ВСЕ найденные фотки для проверки...")

    downloaded_info = []

    for idx, test_image_url in enumerate(unique_images, 1):
        try:
            img_response = requests.get(test_image_url, timeout=10, stream=True)
            img_response.raise_for_status()

            # Определяем расширение
            ext = '.webp' if 'webp' in test_image_url else '.jpg'
            test_file = Path(__file__).parent / f"test_image_{idx}{ext}"

            with open(test_file, 'wb') as f:
                for chunk in img_response.iter_content(chunk_size=8192):
                    f.write(chunk)

            file_size = test_file.stat().st_size / 1024  # KB
            downloaded_info.append({
                'idx': idx,
                'url': test_image_url,
                'file': test_file.name,
                'size': file_size
            })

        except Exception as e:
            print(f"   ❌ Ошибка скачивания #{idx}: {e}")

    # Выводим сводку
    print(f"\n📊 СВОДКА СКАЧАННЫХ ФОТОК:")
    print(f"{'№':<4} {'Размер':<12} {'Файл':<20} {'URL'}")
    print("-" * 100)

    for info in sorted(downloaded_info, key=lambda x: x['size'], reverse=True):
        size_str = f"{info['size']:.1f} KB"
        status = "✅ БОЛЬШАЯ" if info['size'] > 100 else "⚠️ ПРЕВЬЮ" if info['size'] > 20 else "❌ МИНИ"
        print(f"{info['idx']:<4} {size_str:<12} {status:<15} {info['url']}")
else:
    print("\n❌ Фотки не найдены!")

print("\n" + "=" * 70)
print("🏁 ТЕСТ ЗАВЕРШЁН")
