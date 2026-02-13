import requests
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

base_domain = 'https://www.tradeinn.com'
product_id = '141608258'
category_id = '14160'
name = 'asics-Обувь-для-закрытых-кортов-netburner-ballistic-ff-3'

print('\nTesting different URL patterns to find LARGE images:\n')

# Попробуем разные директории и модификаторы
patterns = [
    # Оригинальный паттерн с /f/
    ('Original /f/', f'/f/{category_id}/{product_id}/{name}.webp'),
    ('Original /f/ jpg', f'/f/{category_id}/{product_id}/{name}.jpg'),

    # Попробуем /i/ вместо /f/
    ('/i/ directory', f'/i/{category_id}/{product_id}/{name}.webp'),
    ('/i/ jpg', f'/i/{category_id}/{product_id}/{name}.jpg'),

    # Попробуем /images/
    ('/images/', f'/images/{category_id}/{product_id}/{name}.webp'),

    # Попробуем без категории
    ('No category /f/', f'/f/{product_id}/{name}.webp'),
    ('No category /i/', f'/i/{product_id}/{name}.webp'),

    # Попробуем числовые размеры
    ('/600/', f'/600/{category_id}/{product_id}/{name}.webp'),
    ('/800/', f'/800/{category_id}/{product_id}/{name}.webp'),
    ('/1200/', f'/1200/{category_id}/{product_id}/{name}.webp'),
    ('/1600/', f'/1600/{category_id}/{product_id}/{name}.webp'),

    # Попробуем с модификаторами размера
    ('/f/ _big', f'/f/{category_id}/{product_id}/{name}_big.webp'),
    ('/f/ _full', f'/f/{category_id}/{product_id}/{name}_full.webp'),
    ('/f/ _hd', f'/f/{category_id}/{product_id}/{name}_hd.webp'),
    ('/f/ _original', f'/f/{category_id}/{product_id}/{name}_original.webp'),

    # Попробуем только числовой ID (из JSON-LD)
    ('Numeric 4570158173346', f'/f/{category_id}/{product_id}/4570158173346.webp'),
    ('Numeric jpg', f'/f/{category_id}/{product_id}/4570158173346.jpg'),
]

results = []

for pattern_name, path in patterns:
    url = base_domain + path
    try:
        r = requests.head(url, timeout=5, allow_redirects=True)
        if r.status_code == 200:
            size_kb = int(r.headers.get('content-length', 0)) / 1024
            status = '✅✅✅ BIG!' if size_kb > 100 else '⚠️ PREVIEW' if size_kb > 20 else '❌ TINY'
            results.append((pattern_name, size_kb, status, url))
        else:
            results.append((pattern_name, 0, f'HTTP {r.status_code}', url))
    except Exception as e:
        results.append((pattern_name, 0, f'ERROR', url))

# Sort by size (biggest first)
results.sort(key=lambda x: x[1], reverse=True)

print(f"{'Pattern':<25} | {'Size':<12} | {'Status':<15} | URL")
print('-' * 140)

for pattern_name, size_kb, status, url in results:
    if size_kb > 0:
        print(f'{pattern_name:<25} | {size_kb:8.1f} KB | {status:<15} | {url[:80]}...')
    else:
        print(f'{pattern_name:<25} | {"N/A":>8s}    | {status:<15} | {url[:80]}...')

print('\nDone!')
