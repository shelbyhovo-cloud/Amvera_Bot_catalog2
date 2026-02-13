import requests
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

base_url = 'https://www.tradeinn.com'
product_id = '141608258'
category_id = '14160'
name = 'asics-Обувь-для-закрытых-кортов-netburner-ballistic-ff-3'

print('\nChecking image sizes with different suffixes:\n')

for i in range(1, 10):
    suffix = '' if i == 1 else f'_{i}'
    url = f'{base_url}/f/{category_id}/{product_id}{suffix}/{name}.webp'

    try:
        r = requests.head(url, timeout=5)
        if r.status_code == 200:
            size_kb = int(r.headers.get('content-length', 0)) / 1024
            status = 'BIG' if size_kb > 100 else 'PREVIEW' if size_kb > 20 else 'MINI'
            print(f'{i}. {suffix if suffix else "(base)":10s} | {size_kb:8.1f} KB | {status:10s} | {url[:90]}...')
        else:
            print(f'{i}. {suffix if suffix else "(base)":10s} | HTTP {r.status_code} - not found')
            break
    except Exception as e:
        print(f'Error: {e}')
        break

print('\nDone!')
