"""
Убивает все процессы на указанном порту
"""
import subprocess
import sys
import io
import re

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def kill_process_on_port(port):
    """Убивает все процессы которые используют указанный порт"""
    print(f"🔍 Ищу процессы на порту {port}...")

    try:
        # Находим процессы на порту
        result = subprocess.run(
            f'netstat -ano | findstr :{port}',
            shell=True,
            capture_output=True,
            text=True
        )

        if not result.stdout.strip():
            print(f"✅ Порт {port} свободен")
            return True

        # Извлекаем PID процессов
        pids = set()
        for line in result.stdout.strip().split('\n'):
            parts = line.split()
            if len(parts) >= 5:
                pid = parts[-1]
                if pid.isdigit():
                    pids.add(pid)

        if not pids:
            print(f"✅ Порт {port} свободен")
            return True

        # Убиваем каждый процесс
        killed_count = 0
        for pid in pids:
            try:
                subprocess.run(
                    f'taskkill /PID {pid} /F',
                    shell=True,
                    capture_output=True,
                    check=True
                )
                print(f"❌ Убил процесс PID {pid}")
                killed_count += 1
            except subprocess.CalledProcessError:
                print(f"⚠️ Не удалось убить процесс PID {pid}")

        print(f"\n✅ Порт {port} освобождён! Убито процессов: {killed_count}")
        return True

    except Exception as e:
        print(f"❌ Ошибка: {e}")
        return False

if __name__ == "__main__":
    port = sys.argv[1] if len(sys.argv) > 1 else "8080"
    kill_process_on_port(port)
