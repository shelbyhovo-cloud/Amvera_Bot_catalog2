"""
Полный тест Serveo подключения
"""
import subprocess
import sys
import io
import time
import re

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

print("=" * 60)
print("🧪 ТЕСТ SERVEO ПОДКЛЮЧЕНИЯ")
print("=" * 60)
print()

# Тест 1: Проверка SSH клиента
print("1️⃣ Проверяю SSH клиент...")
try:
    result = subprocess.run(['ssh', '-V'], capture_output=True, text=True, timeout=5)
    version = result.stderr.strip() if result.stderr else result.stdout.strip()
    print(f"   ✅ SSH найден: {version[:50]}")
except FileNotFoundError:
    print("   ❌ SSH не найден!")
    print("   Установи: Параметры → Приложения → Дополнительные компоненты → OpenSSH Client")
    sys.exit(1)
except Exception as e:
    print(f"   ⚠️  Ошибка: {e}")

print()

# Тест 2: Проверка сетевого подключения к serveo.net
print("2️⃣ Проверяю доступность serveo.net...")
try:
    # Пробуем ping
    result = subprocess.run(['ping', '-n', '2', 'serveo.net'],
                          capture_output=True, text=True, timeout=10)
    if 'TTL=' in result.stdout or 'ttl=' in result.stdout:
        print("   ✅ Serveo.net доступен")
    else:
        print("   ⚠️  Ping не прошёл, но это не критично")
except Exception as e:
    print(f"   ⚠️  Не удалось проверить ping: {e}")

print()

# Тест 3: Попытка подключения к Serveo
print("3️⃣ Пробую подключиться к Serveo...")
print("   (Это может занять до 15 секунд)")
print()

try:
    # Запускаем SSH туннель
    process = subprocess.Popen(
        ['ssh', '-o', 'StrictHostKeyChecking=no',
         '-o', 'ConnectTimeout=10',
         '-o', 'ServerAliveInterval=30',
         '-R', '80:localhost:8080', 'serveo.net'],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        stdin=subprocess.PIPE,
        text=True,
        bufsize=1
    )

    url_found = None
    start_time = time.time()
    timeout = 15  # 15 секунд

    print("   Читаю вывод от Serveo:")
    print("   " + "-" * 50)

    while time.time() - start_time < timeout:
        # Проверяем что процесс жив
        if process.poll() is not None:
            print("\n   ⚠️  Процесс завершился преждевременно")
            stderr = process.stdout.read() if process.stdout else ""
            if stderr:
                print(f"   Вывод: {stderr[:300]}")
            break

        line = process.stdout.readline()
        if line:
            line_clean = line.strip()
            if line_clean:
                print(f"   {line_clean}")

            # Ищем URL
            match = re.search(r'https://[a-zA-Z0-9\-]+\.serveo\.net', line)
            if match:
                url_found = match.group(0)
                print()
                print(f"   ✅ ПОЛУЧЕН URL: {url_found}")
                break

    if url_found:
        print()
        print("=" * 60)
        print("✅ ТЕСТ ПРОЙДЕН!")
        print("=" * 60)
        print(f"\nPublic URL: {url_found}")
        print("\nServeo работает! Можешь запускать бота.")
        print()
    else:
        print()
        print("=" * 60)
        print("❌ ТЕСТ НЕ ПРОЙДЕН")
        print("=" * 60)
        print()
        print("Serveo не ответил за 15 секунд.")
        print()
        print("Возможные причины:")
        print("  1. Serveo сервер перегружен или недоступен")
        print("  2. Файрвол блокирует SSH (порт 22)")
        print("  3. Проблемы с интернет-соединением")
        print()
        print("Что попробовать:")
        print("  1. Подожди 5 минут и попробуй снова")
        print("  2. Используй VPN если есть")
        print("  3. Попробуй альтернативу: LocalTunnel")
        print("     npx localtunnel --port 8080")
        print()

    # Останавливаем процесс
    try:
        process.kill()
    except:
        pass

except KeyboardInterrupt:
    print("\n\n🛑 Остановлено пользователем")
    try:
        process.kill()
    except:
        pass
except Exception as e:
    print(f"\n❌ Ошибка теста: {e}")
    import traceback
    traceback.print_exc()

print()
print("=" * 60)
input("Нажми Enter чтобы закрыть...")
