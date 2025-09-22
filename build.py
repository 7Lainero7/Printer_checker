import os
import subprocess
import sys
import shutil
from pathlib import Path
from PIL import Image

def convert_png_to_ico(png_path, ico_path):
    """Конвертирует PNG в ICO если нужно"""
    if os.path.exists(png_path) and not os.path.exists(ico_path):
        try:
            img = Image.open(png_path)
            # Сохраняем с разными размерами для иконки
            img.save(ico_path, format='ICO', sizes=[(16,16), (32,32), (48,48), (64,64)])
            print(f"✅ Иконка создана: {ico_path}")
            return True
        except Exception as e:
            print(f"❌ Ошибка конвертации иконки: {e}")
            return False
    return True

def build_exe():
    """Сборка EXE файла с включением DLL и иконки"""
    
    # Конвертируем PNG в ICO если нужно
    ico_path = "printer_icon.ico"
    if not os.path.exists(ico_path) and os.path.exists("printer.png"):
        if not convert_png_to_ico("printer.png", ico_path):
            print("❌ Продолжаем без иконки")
            ico_path = None
    elif not os.path.exists(ico_path):
        print("❌ Иконка не найдена, продолжаем без иконки")
        ico_path = None
    
    # Находим путь к pylibdmtx
    try:
        import pylibdmtx
        lib_path = Path(pylibdmtx.__file__).parent
        print(f"Путь к pylibdmtx: {lib_path}")
        
        # Ищем DLL файлы
        dll_files = []
        for root, dirs, files in os.walk(lib_path):
            for file in files:
                if file.endswith('.dll'):
                    dll_files.append(os.path.join(root, file))
        
        if not dll_files:
            print("❌ DLL файлы не найдены!")
            return False
            
        print(f"Найдено DLL файлов: {len(dll_files)}")
        for dll in dll_files:
            print(f"  - {dll}")
            
    except ImportError:
        print("❌ pylibdmtx не установлен")
        return False
    
    # Базовая команда для сборки
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name=DataMatrixPrinter",
        "--clean",
        "--add-data", f"{lib_path}{os.pathsep}pylibdmtx",
        "--hidden-import=win32print",
        "--hidden-import=win32ui", 
        "--hidden-import=win32con",
        "--hidden-import=sqlite3",
        "--hidden-import=PIL",
        "--hidden-import=PIL.Image",
        "--hidden-import=PIL.ImageWin",
        "--hidden-import=PyQt5",
        "--hidden-import=PyQt5.QtCore",
        "--hidden-import=PyQt5.QtGui",
        "--hidden-import=PyQt5.QtWidgets",
    ]
    
    # Добавляем иконку если есть
    if ico_path and os.path.exists(ico_path):
        cmd.extend(["--icon", ico_path])
        print(f"🎨 Используем иконку: {ico_path}")
    
    cmd.append("main.py")
    
    print("🔨 Начинаем сборку EXE файла...")
    print(f"Команда: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        print("✅ Сборка завершена успешно!")
        
        # Проверяем результат
        exe_path = os.path.join('dist', 'DataMatrixPrinter.exe')
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"📊 Размер файла: {size_mb:.1f} MB")
            print(f"📁 EXE файл: {exe_path}")
        
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Ошибка сборки:")
        print(f"STDERR: {e.stderr}")
        
        # Пробуем без DLL если не получилось
        print("🔄 Пробуем сборку без явного включения DLL...")
        cmd_simple = [c for c in cmd if '--add-data' not in c]
        try:
            result = subprocess.run(cmd_simple, capture_output=True, text=True, check=True)
            print("✅ Сборка без DLL завершена успешно!")
            return True
        except:
            return False

if __name__ == "__main__":
    print("🚀 Сборка DataMatrix Printer с иконкой и DLL")
    print("=" * 50)
    
    # Очищаем предыдущие сборки
    if os.path.exists("build"):
        shutil.rmtree("build")
    if os.path.exists("dist"):
        shutil.rmtree("dist")
    if os.path.exists("DataMatrixPrinter.spec"):
        os.remove("DataMatrixPrinter.spec")
    
    success = build_exe()
    
    if success:
        print("\n🎉 Сборка завершена!")
        print("📋 Файл готов: dist/DataMatrixPrinter.exe")
    else:
        print("\n💥 Ошибка сборки!")
    
    input("\nНажмите Enter...")