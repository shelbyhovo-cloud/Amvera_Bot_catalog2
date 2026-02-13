"""
Простой способ получить публичный URL через Serveo (не нужен ngrok)
"""
import subprocess
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

print("🌐 Serveo - бесплатная альтернатива ngrok")
print("=" * 50)
print("\n📝 ИНСТРУКЦИЯ:")
print("1. Открой новый терминал")
print("2. Запусти команду:")
print('\n   ssh -R 80:localhost:8080 serveo.net\n')
print("3. Скопируй полученный URL (будет что-то вроде: https://abc123.serveo.net)")
print("4. Вставь этот URL в mini_app.py:")
print('   MODE = "manual"')
print('   MANUAL_WEBAPP_URL = "твой_url_из_serveo"')
print("\n5. Запусти бота: python mini_app.py")
print("\n⚠️ ВАЖНО: Не закрывай окно с ssh командой!")
print("=" * 50)

input("\nНажми Enter чтобы закрыть...")
