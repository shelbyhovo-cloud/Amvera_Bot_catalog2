"""
ПОЛНОСТЬЮ АВТОМАТИЧЕСКИЙ ЗАПУСК БОТА
Пробует Serveo, если не работает - запускает локально
"""
import subprocess
import sys
import io
import time
import re
import webbrowser
from pathlib import Path
from threading import Thread

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

print("=" * 60)
print("🚀 АВТОМАТИЧЕСКИЙ ЗАПУСК MINI APP")
print("=" * 60)
print()

def try_serveo():
    """Пробует запустить Serveo туннель"""
    print("1️⃣ Пробую автоматический туннель (Serveo)...")

    try:
        # Запускаем SSH туннель
        process = subprocess.Popen(
            ['ssh', '-o', 'StrictHostKeyChecking=no',
             '-o', 'ConnectTimeout=5',
             '-R', '80:localhost:8080', 'serveo.net'],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.PIPE,
            text=True,
            bufsize=1
        )

        # Ждем URL максимум 10 секунд
        url = None
        for i in range(20):
            if process.poll() is not None:
                break

            line = process.stdout.readline()
            if line:
                match = re.search(r'https://[a-zA-Z0-9\-]+\.serveo\.net', line)
                if match:
                    url = match.group(0)
                    break
            time.sleep(0.5)

        if url:
            print(f"   ✅ Serveo работает: {url}")

            # Обновляем настройки
            config_path = Path(__file__).parent / "mini_app.py"
            with open(config_path, 'r', encoding='utf-8') as f:
                content = f.read()

            content = re.sub(r'MODE = "[^"]*"', 'MODE = "manual"', content)
            content = re.sub(r'MANUAL_WEBAPP_URL = "[^"]*"', f'MANUAL_WEBAPP_URL = "{url}"', content)

            with open(config_path, 'w', encoding='utf-8') as f:
                f.write(content)

            print("   ✅ Настройки обновлены")
            return process, url
        else:
            print("   ⚠️  Serveo не ответил за 10 секунд")
            try:
                process.kill()
            except:
                pass
            return None, None

    except Exception as e:
        print(f"   ❌ Serveo не работает: {e}")
        return None, None

def setup_local_mode():
    """Настраивает локальный режим"""
    print("2️⃣ Настраиваю локальный режим...")

    config_path = Path(__file__).parent / "mini_app.py"
    with open(config_path, 'r', encoding='utf-8') as f:
        content = f.read()

    content = re.sub(r'MODE = "[^"]*"', 'MODE = "manual"', content)
    content = re.sub(r'MANUAL_WEBAPP_URL = "[^"]*"', 'MANUAL_WEBAPP_URL = "http://localhost:8080"', content)

    with open(config_path, 'w', encoding='utf-8') as f:
        f.write(content)

    print("   ✅ Локальный режим активирован")
    print("   📌 Веб-интерфейс: http://localhost:8080")
    print()

def start_bot():
    """Запускает бота"""
    print("3️⃣ Запускаю бота...")
    print("=" * 60)
    print()

    bot_script = Path(__file__).parent / "mini_app.py"

    try:
        subprocess.run([sys.executable, str(bot_script)])
    except KeyboardInterrupt:
        print("\n🛑 Остановлено пользователем")

# Главная логика
if __name__ == "__main__":
    try:
        # Пробуем Serveo
        serveo_process, serveo_url = try_serveo()

        if not serveo_url:
            # Serveo не работает - локальный режим
            print()
            print("⚠️  Serveo недоступен - переключаюсь на локальный режим")
            print()
            setup_local_mode()
            print("⚠️  ВНИМАНИЕ:")
            print("   - Telegram Mini App НЕ БУДЕТ РАБОТАТЬ")
            print("   - Можно только открыть веб-интерфейс в браузере")
            print("   - Для полной работы нужен публичный HTTPS URL")
            print()

            # Открываем браузер через 3 секунды
            def open_browser():
                time.sleep(3)
                try:
                    webbrowser.open('http://localhost:8080')
                except:
                    pass

            Thread(target=open_browser, daemon=True).start()
        else:
            print()
            print("✅ Всё готово! Бот работает через Serveo")
            print()

        # Запускаем бота
        start_bot()

        # Останавливаем Serveo если был запущен
        if serveo_process:
            try:
                serveo_process.kill()
            except:
                pass

    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()

    input("\nНажми Enter чтобы закрыть...")
