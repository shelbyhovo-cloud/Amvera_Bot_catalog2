"""
Альтернативный способ запуска с localtunnel вместо ngrok
"""
import subprocess
import sys
import time
import requests
from pathlib import Path

print("🌐 Устанавливаю localtunnel...")
try:
    # Проверяем наличие npm
    subprocess.run(["npm", "--version"], check=True, capture_output=True)

    # Устанавливаем localtunnel
    subprocess.run(["npm", "install", "-g", "localtunnel"], check=True)
    print("✅ Localtunnel установлен!\n")

    print("🚀 Запускаю localtunnel на порту 8080...")
    print("⚠️ Не закрывай это окно!\n")

    # Запускаем localtunnel
    process = subprocess.Popen(
        ["lt", "--port", "8080"],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1
    )

    # Ждем URL
    for line in process.stdout:
        print(line.strip())
        if "your url is:" in line.lower():
            url = line.split(":")[-1].strip()
            print(f"\n✅ Публичный URL: {url}")
            print(f"\n📝 Скопируй этот URL и вставь в mini_app.py:")
            print(f'   MODE = "manual"')
            print(f'   MANUAL_WEBAPP_URL = "{url}"')
            print(f"\nЗатем перезапусти бота в другом окне терминала\n")
            break

    # Держим процесс запущенным
    process.wait()

except FileNotFoundError:
    print("❌ NPM не найден!")
    print("\n📝 Установи Node.js:")
    print("   1. Скачай с https://nodejs.org/")
    print("   2. Установи Node.js")
    print("   3. Перезапусти этот скрипт\n")

except subprocess.CalledProcessError as e:
    print(f"❌ Ошибка: {e}")
except KeyboardInterrupt:
    print("\n🛑 Остановлено")
