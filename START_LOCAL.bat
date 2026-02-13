@echo off
chcp 65001 >nul
title Telegram Mini App - Локальный режим

echo ============================================================
echo 🏠 ЛОКАЛЬНЫЙ РЕЖИМ (БЕЗ ТУННЕЛЯ)
echo ============================================================
echo.

cd /d "%~dp0"

echo ⚙️ Настраиваю локальный режим...
python -c "import re; content = open('mini_app.py', 'r', encoding='utf-8').read(); content = re.sub(r'MODE = \"auto\"', 'MODE = \"manual\"', content); content = re.sub(r'MANUAL_WEBAPP_URL = \"[^\"]*\"', 'MANUAL_WEBAPP_URL = \"http://localhost:8080\"', content); open('mini_app.py', 'w', encoding='utf-8').write(content)"

echo ✅ Локальный режим активирован
echo.
echo 🚀 Запускаю бота...
echo.
echo 📌 После запуска открой в браузере:
echo    http://localhost:8080
echo.
echo ⚠️ ВНИМАНИЕ: Telegram Mini App не будет работать!
echo    Это только для просмотра веб-интерфейса.
echo.

python mini_app.py

pause
