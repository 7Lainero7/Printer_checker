# build_fixed.py
import os
import subprocess
import sys
import shutil
from pathlib import Path

def build_exe():
    """Сборка EXE файла с абсолютным путем к иконке"""
    
    # Получаем абсолютный путь к иконке
    ico_path = os.path.abspath("printer_icon.ico")
    
    if not os.path.exists(ico_path):
        print("❌ Иконка не найдена, собираем без иконки")
        ico_path = None
    
    # Команда для сборки
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name=QRPrinterCheck",
        "--distpath=dist",
        "--workpath=build",
        "--specpath=build",
        "--clean",
        "--hidden-import=win32print",
        "--hidden-import=win32ui", 
        "--hidden-import=qrcode",
        "--hidden-import=PIL",
        "--hidden-import=PyQt5",
        "--hidden-import=PyQt5.QtCore",
        "--hidden-import=PyQt5.QtGui",
        "--hidden-import=PyQt5.QtWidgets",
    ]
    
    # Добавляем иконку с абсолютным путем
    if ico_path:
        cmd.extend(["--icon", ico_path])
        print(f"🎨 Используем иконку: {ico_path}")
    
    cmd.append("main.py")
    
    print("🔨 Начинаем сборку EXE файла...")
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        print("✅ Сборка завершена успешно!")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Ошибка сборки:")
        print(f"STDERR: {e.stderr}")
        
        # Пробуем без иконки
        if ico_path:
            print("🔄 Пробуем сборку без иконки...")
            cmd_no_icon = [c for c in cmd if c != '--icon' and c != ico_path]
            try:
                result = subprocess.run(cmd_no_icon, capture_output=True, text=True, check=True)
                print("✅ Сборка без иконки завершена успешно!")
                return True
            except:
                return False
        return False

if __name__ == "__main__":
    print("🚀 Запуск сборки с абсолютным путем к иконке")
    
    # Очищаем предыдущие сборки
    if os.path.exists("build"):
        shutil.rmtree("build")
    if os.path.exists("dist"):
        shutil.rmtree("dist")
    if os.path.exists("QRPrinterCheck.spec"):
        os.remove("QRPrinterCheck.spec")
    
    success = build_exe()
    
    if success:
        print("\n🎉 Сборка завершена!")
        print("📋 Файл: dist/QRPrinterCheck.exe")
    else:
        print("\n💥 Ошибка сборки!")
    
    input("\nНажмите Enter...")