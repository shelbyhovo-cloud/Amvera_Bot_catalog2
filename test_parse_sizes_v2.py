"""
Тестовый парсинг размеров - альтернативные методы
"""
import sys
import io

# Фикс кодировки для Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import requests
import re
import json

url = "https://www.tradeinn.com/volleyball/ru/asics-%D0%9E%D0%B1%D1%83%D0%B2%D1%8C-%D0%B4%D0%BB%D1%8F-%D0%B7%D0%B0%D0%BA%D1%80%D1%8B%D1%82%D1%8B%D1%85-%D0%BA%D0%BE%D1%80%D1%82%D0%BE%D0%B2-netburner-ballistic-ff-3/141608258/p"

print("=" * 80)
print("🧪 ТЕСТОВЫЙ ПАРСИНГ РАЗМЕРОВ V2")
print("=" * 80)

print("\n🌐 Загружаю страницу...")
response = requests.get(url, timeout=10)
html = response.text
print(f"✅ Загружено {len(html)} байт\n")

sizes = []

# МЕТОД 1: Ищем JavaScript переменную productSizes или similar
print("🔍 МЕТОД 1: Ищу JavaScript переменные с размерами...")
js_patterns = [
    r'sizes["\']?\s*:\s*\[([^\]]+)\]',
    r'productSizes\s*=\s*\[([^\]]+)\]',
    r'availableSizes\s*=\s*\[([^\]]+)\]',
    r'"sizeList"\s*:\s*\[([^\]]+)\]',
]

for pattern in js_patterns:
    matches = re.search(pattern, html, re.IGNORECASE)
    if matches:
        print(f"   ✅ Найдено совпадение: {pattern[:30]}...")
        print(f"      Данные: {matches.group(1)[:100]}...")
        try:
            # Пробуем распарсить как JSON
            sizes_data = json.loads('[' + matches.group(1) + ']')
            for size in sizes_data:
                if isinstance(size, (str, dict)):
                    if isinstance(size, dict):
                        size_val = size.get('size') or size.get('value') or size.get('name')
                    else:
                        size_val = size

                    if size_val and str(size_val).strip():
                        sizes.append(str(size_val).strip())
                        print(f"      📏 Размер: {size_val}")
        except:
            pass

# МЕТОД 2: Ищем data-атрибуты с размерами
if not sizes:
    print("\n🔍 МЕТОД 2: Ищу data-атрибуты...")
    data_patterns = [
        r'data-size="([^"]+)"',
        r'data-variant-size="([^"]+)"',
        r'data-dimension="([^"]+)"',
    ]

    for pattern in data_patterns:
        matches = re.findall(pattern, html, re.IGNORECASE)
        if matches:
            print(f"   ✅ Найдено {len(matches)} совпадений по pattern: {pattern}")
            for match in matches[:10]:  # Первые 10
                if match.strip() and len(match) <= 15:
                    sizes.append(match.strip())
                    print(f"      📏 Размер: {match}")

# МЕТОД 3: Ищем в HTML select/option элементах
if not sizes:
    print("\n🔍 МЕТОД 3: Ищу в select/option...")
    option_matches = re.findall(r'<option[^>]*>([^<]*(?:\d{2}[.,\s]?\d?/?\d?)[^<]*)</option>', html, re.IGNORECASE)
    print(f"   Найдено option элементов: {len(option_matches)}")

    for option in option_matches[:20]:
        # Ищем числа от 35 до 50 (размеры обуви)
        size_match = re.search(r'(\d{2}(?:[.,\s]?\d)?(?:\s*1/2)?)', option)
        if size_match:
            size = size_match.group(1).strip()
            try:
                size_num = float(size.replace(',', '.').replace(' ', ''))
                if 35 <= size_num <= 50:
                    if size not in sizes:
                        sizes.append(size)
                        print(f"      📏 Размер: {size} (из: {option[:50]})")
            except:
                pass

# МЕТОД 4: Ищем паттерны "EU 42", "Size 42" и т.д.
if not sizes:
    print("\n🔍 МЕТОД 4: Ищу текстовые паттерны размеров...")
    text_patterns = [
        r'(?:EU|Size|Размер)\s+(\d{2}(?:\s*1/2)?)',
        r'size["\']?\s*:\s*["\'](\d{2}(?:\s*1/2)?)["\']',
    ]

    for pattern in text_patterns:
        matches = re.findall(pattern, html, re.IGNORECASE)
        if matches:
            print(f"   ✅ Найдено {len(matches)} размеров")
            for match in set(matches):
                if match not in sizes:
                    sizes.append(match)
                    print(f"      📏 Размер: {match}")

print("\n" + "=" * 80)
print("📊 РЕЗУЛЬТАТ")
print("=" * 80)

if sizes:
    # Убираем дубликаты
    sizes = list(dict.fromkeys(sizes))

    # Сортируем
    def parse_size(s):
        if '1/2' in s:
            base = float(s.replace('1/2', '').strip())
            return base + 0.5
        try:
            return float(s.replace(',', '.').replace(' ', ''))
        except:
            return 999

    sizes = sorted(set(sizes), key=parse_size)

    print(f"\n✅ Найдено размеров: {len(sizes)}")
    print(f"\n📏 Размеры: {', '.join(sizes)}")
    print(f"\n💾 Для Excel: {', '.join(sizes)}")
else:
    print("\n❌ Размеры не найдены!")

print("\n" + "=" * 80)
