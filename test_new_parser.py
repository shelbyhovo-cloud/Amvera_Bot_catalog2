"""
Тестируем обновлённый парсер
"""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import requests
import re
from pathlib import Path

def parse_tradeinn_product_test(url):
    """Тестовая версия парсера"""
    if '?' in url:
        url = url.split('?')[0]

    if '/en/' in url:
        url = url.replace('/en/', '/ru/')

    # Извлекаем product_id из URL
    url_product_id = None
    product_id_match = re.search(r'/(\d+)/p/?$', url)
    if product_id_match:
        url_product_id = product_id_match.group(1)

    response = requests.get(url, timeout=10)
    response.raise_for_status()
    html = response.text

    image_urls = []

    # МЕТОД 1: Галерея с data-fancybox
    print("\n🔍 МЕТОД 1: Ищем галерею data-fancybox...")
    gallery_links = re.findall(r'data-fancybox="gallery"[^>]*href="([^"]+)"', html, re.IGNORECASE)
    if gallery_links:
        print(f"   ✅ Найдено {len(gallery_links)} изображений в галерее")
        for i, link in enumerate(gallery_links, 1):
            if link.startswith('/'):
                link = 'https://www.tradeinn.com' + link
            if link not in image_urls:
                image_urls.append(link)
            print(f"   {i}. {link}")
    else:
        print("   ❌ Галерея не найдена")

    # МЕТОД 2: Паттерн с суффиксами
    if not image_urls and url_product_id:
        print("\n🔍 МЕТОД 2: Ищем по паттерну с суффиксами...")
        category_match = re.search(r'/(\d+)/\d+/p', url)
        if category_match:
            category_id = category_match.group(1)
            pattern = rf'/f/{category_id}/{url_product_id}(?:_\d+)?/[^"\']+\.(?:jpg|jpeg|png|webp)'
            found_images = re.findall(pattern, html, re.IGNORECASE)

            print(f"   Найдено {len(found_images)} изображений")
            for img in found_images:
                full_url = 'https://www.tradeinn.com' + img if img.startswith('/') else img
                if full_url not in image_urls:
                    image_urls.append(full_url)

    print(f"\n📊 ИТОГО найдено уникальных изображений: {len(image_urls)}")
    return image_urls


# Тестируем
url = "https://www.tradeinn.com/volleyball/ru/asics-Обувь-для-закрытых-кортов-netburner-ballistic-ff-3/141608258/p"

print("=" * 80)
print("🧪 ТЕСТИРУЮ ОБНОВЛЁННЫЙ ПАРСЕР")
print("=" * 80)
print(f"URL: {url}\n")

images = parse_tradeinn_product_test(url)

print("\n" + "=" * 80)
print("✅ ТЕСТ ЗАВЕРШЁН")
print("=" * 80)
