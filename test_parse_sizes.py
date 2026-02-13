"""
Тестовый парсинг размеров с TradeInn
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
print("🧪 ТЕСТОВЫЙ ПАРСИНГ РАЗМЕРОВ")
print("=" * 80)
print(f"\n📝 URL: {url}\n")

print("🌐 Загружаю страницу...")
response = requests.get(url, timeout=10)
html = response.text
print(f"✅ Страница загружена ({len(html)} байт)\n")

# Парсим название
name_match = re.search(r'<h1[^>]*>([^<]+)</h1>', html, re.IGNORECASE)
name = name_match.group(1).strip() if name_match else "Без названия"
print(f"📦 Товар: {name}\n")

# Парсим размеры
sizes = []

print("🔍 Ищу размеры в JSON-LD...")
json_ld_pattern = r'<script type="application/ld\+json">(.*?)</script>'
json_ld_matches = re.findall(json_ld_pattern, html, re.DOTALL)

print(f"   Найдено JSON-LD блоков: {len(json_ld_matches)}")

for idx, json_str in enumerate(json_ld_matches):
    try:
        data = json.loads(json_str)
        print(f"\n   📄 Блок {idx + 1}:")
        print(f"      @type: {data.get('@type') if isinstance(data, dict) else 'N/A'}")

        if isinstance(data, dict) and data.get('@type') == 'Product':
            print(f"      ✅ Это Product!")
            print(f"      Название: {data.get('name', 'N/A')}")

            # Показываем все ключи
            print(f"      Ключи: {', '.join(data.keys())}")

            variants = data.get('hasVariant', [])
            print(f"      hasVariant: {len(variants)} вариантов")

            if variants:
                print(f"\n      🔍 Детали вариантов:")
                for v_idx, variant in enumerate(variants[:3]):  # Показываем первые 3
                    print(f"         Вариант {v_idx + 1}:")
                    print(f"            name: {variant.get('name', 'N/A')}")
                    print(f"            sku: {variant.get('sku', 'N/A')}")
                    if variant.get('additionalProperty'):
                        for prop in variant.get('additionalProperty', []):
                            print(f"            {prop.get('name')}: {prop.get('value')}")

                    variant_name = variant.get('name', '')
                    # Пример: "EU 42 1/2" или "EU 44"
                    size_match = re.search(r'EU\s+(\d+(?:\s*1/2)?)', variant_name)
                    if size_match:
                        size = size_match.group(1).strip()
                        if size not in sizes:
                            sizes.append(size)
                            print(f"            ✅ Извлечен размер: {size}")
        else:
            print(f"      ⚠️  Не Product, тип: {data.get('@type') if isinstance(data, dict) else type(data)}")

    except Exception as e:
        print(f"   ❌ Блок {idx + 1}: Ошибка парсинга - {e}")
        import traceback
        traceback.print_exc()

print("\n" + "=" * 80)
print("📊 РЕЗУЛЬТАТ")
print("=" * 80)

if sizes:
    # Сортируем размеры
    def parse_size(s):
        if '1/2' in s:
            base = float(s.replace('1/2', '').strip())
            return base + 0.5
        try:
            return float(s.replace(',', '.'))
        except:
            return 999

    sizes = sorted(set(sizes), key=parse_size)

    print(f"\n✅ Найдено размеров: {len(sizes)}")
    print(f"\n📏 Размеры: {', '.join(sizes)}")
    print(f"\n💾 Для Excel: {', '.join(sizes)}")
else:
    print("\n❌ Размеры не найдены!")
    print("\n💡 Возможные причины:")
    print("   • Размеры загружаются динамически (JavaScript)")
    print("   • Структура страницы изменилась")
    print("   • Товар снят с продажи")

print("\n" + "=" * 80)
input("\nНажми Enter для выхода...")
