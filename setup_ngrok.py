"""
Автоматическая настройка ngrok
"""
import subprocess
import sys
import io
from pathlib import Path

# Фикс кодировки
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Путь к ngrok
ngrok_path = Path(__file__).parent / "ngrok_bin" / "ngrok.exe"
authtoken = "39ckmuBIwfRHsy7zxQL5WngTVPE_5HeeiGpyQqN9CbR2wpMJN"

print("🔧 Настраиваю ngrok...")
print(f"📁 Путь: {ngrok_path}\n")

try:
    # Настраиваем authtoken
    result = subprocess.run(
        [str(ngrok_path), "config", "add-authtoken", authtoken],
        capture_output=True,
        text=True,
        check=True
    )

    print("✅ Ngrok успешно настроен!")
    print("\n📝 Теперь запусти бота:")
    print("   python mini_app.py\n")

except subprocess.CalledProcessError as e:
    print(f"❌ Ошибка настройки: {e}")
    print(f"Вывод: {e.stderr}")
    sys.exit(1)
except FileNotFoundError:
    print(f"❌ Ngrok не найден по пути: {ngrok_path}")
    sys.exit(1)
