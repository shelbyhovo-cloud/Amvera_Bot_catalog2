"""
Автоматический запуск Mini App с Serveo (альтернатива ngrok)
Serveo - бесплатный сервис туннелирования без регистрации
"""
import subprocess
import sys
import os
import io
import time
import re
from pathlib import Path
from threading import Thread

# Фикс кодировки
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

print("=" * 60)
print("🌐 ЗАПУСК MINI APP С SERVEO")
print("=" * 60)
print()

# Порт на котором работает бот
PORT = 8080
serveo_url = None
serveo_process = None


def start_serveo():
    """Запускает Serveo туннель и парсит URL"""
    global serveo_url, serveo_process

    print("🔧 Запускаю Serveo туннель...")
    print(f"   Локальный порт: {PORT}")
    print()

    try:
        # Запускаем SSH туннель к Serveo
        serveo_process = subprocess.Popen(
            ['ssh', '-o', 'StrictHostKeyChecking=no', '-R', f'80:localhost:{PORT}', 'serveo.net'],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.PIPE,
            text=True,
            bufsize=1,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
        )

        # Читаем вывод и ищем URL
        print("⏳ Получаю публичный URL от Serveo...")
        for line in serveo_process.stdout:
            print(f"   Serveo: {line.strip()}")

            # Ищем URL в формате: Forwarding HTTP traffic from https://xxxxx.serveo.net
            match = re.search(r'https://[a-zA-Z0-9\-]+\.serveo\.net', line)
            if match:
                serveo_url = match.group(0)
                print()
                print("✅ Serveo туннель активен!")
                print(f"🌍 Публичный URL: {serveo_url}")
                print()
                break

            # Альтернативный формат
            if 'serveo.net' in line.lower():
                # Попробуем извлечь любой URL
                urls = re.findall(r'https://[^\s]+', line)
                if urls:
                    serveo_url = urls[0]
                    print()
                    print("✅ Serveo туннель активен!")
                    print(f"🌍 Публичный URL: {serveo_url}")
                    print()
                    break

        if not serveo_url:
            print("❌ Не удалось получить URL от Serveo")
            return False

        # Продолжаем читать вывод в фоне
        def read_output():
            for line in serveo_process.stdout:
                pass  # Просто читаем чтобы не блокировался буфер

        Thread(target=read_output, daemon=True).start()
        return True

    except FileNotFoundError:
        print("❌ SSH не найден!")
        print()
        print("📝 Установи SSH клиент:")
        print("   Windows 10/11: Уже должен быть установлен")
        print("   Если нет, включи в: Параметры → Приложения → Дополнительные компоненты → OpenSSH Client")
        print()
        return False
    except Exception as e:
        print(f"❌ Ошибка запуска Serveo: {e}")
        return False


def start_bot():
    """Запускает бота в ручном режиме с Serveo URL"""
    print("=" * 60)
    print("🤖 ЗАПУСК TELEGRAM БОТА")
    print("=" * 60)
    print()

    # Устанавливаем переменные окружения для бота
    os.environ['WEBAPP_URL'] = serveo_url
    os.environ['MODE'] = 'manual'

    # Запускаем бота как subprocess
    try:
        bot_script = Path(__file__).parent / "mini_app.py"

        # Модифицируем mini_app.py временно
        with open(bot_script, 'r', encoding='utf-8') as f:
            code = f.read()

        # Заменяем настройки
        code = re.sub(r'MODE = "[^"]*"', f'MODE = "manual"', code)
        code = re.sub(r'MANUAL_WEBAPP_URL = "[^"]*"', f'MANUAL_WEBAPP_URL = "{serveo_url}"', code)

        # Сохраняем временно
        temp_script = Path(__file__).parent / "mini_app_temp.py"
        with open(temp_script, 'w', encoding='utf-8') as f:
            f.write(code)

        # Запускаем временный скрипт
        bot_process = subprocess.Popen(
            [sys.executable, str(temp_script)],
            stdout=sys.stdout,
            stderr=sys.stderr
        )

        # Ждем завершения
        bot_process.wait()

        # Удаляем временный файл
        if temp_script.exists():
            temp_script.unlink()

    except KeyboardInterrupt:
        print("\n🛑 Останавливаю бота...")
    except Exception as e:
        print(f"❌ Ошибка запуска бота: {e}")
    finally:
        # Останавливаем Serveo
        if serveo_process:
            print("🛑 Останавливаю Serveo туннель...")
            serveo_process.terminate()


if __name__ == "__main__":
    try:
        # Запускаем Serveo
        if start_serveo():
            time.sleep(1)  # Даем время туннелю стабилизироваться

            # Запускаем бота
            start_bot()
        else:
            print()
            print("💡 АЛЬТЕРНАТИВНЫЙ СПОСОБ:")
            print("   1. Открой новый терминал")
            print(f"   2. Запусти: ssh -R 80:localhost:{PORT} serveo.net")
            print("   3. Скопируй полученный URL")
            print("   4. В mini_app.py измени:")
            print('      MODE = "manual"')
            print('      MANUAL_WEBAPP_URL = "твой_url"')
            print("   5. Запусти: python mini_app.py")

    except KeyboardInterrupt:
        print("\n\n🛑 Остановлено пользователем")
    finally:
        if serveo_process:
            serveo_process.terminate()
