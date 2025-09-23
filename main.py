import sys
import sqlite3
import tempfile
import os
import win32print
import win32ui
import win32con
from datetime import datetime
from PIL import Image, ImageWin
from PyQt5 import QtWidgets, QtCore

try:
    from pylibdmtx.pylibdmtx import encode
except ImportError:
    print("Установите библиотеку: pip install pylibdmtx")
    sys.exit(1)

class PrinterApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QtCore.QSettings("PrinterCheck", "DataMatrixPrinter")
        self.current_printer_info = None
        self.init_ui()
        self.init_database()
        self.find_printers()
        self.load_settings()
        
    def init_ui(self):
        self.setWindowTitle("DataMatrix Printer - Честный знак (GS1)")
        self.setGeometry(100, 100, 900, 800)
        
        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QtWidgets.QVBoxLayout()
        central_widget.setLayout(layout)
        
        # Форма выбора принтера
        printer_group = QtWidgets.QGroupBox("Настройки печати")
        printer_layout = QtWidgets.QGridLayout()
        
        printer_layout.addWidget(QtWidgets.QLabel("Принтер:"), 0, 0)
        self.printer_combo = QtWidgets.QComboBox()
        self.printer_combo.currentIndexChanged.connect(self.on_printer_changed)
        printer_layout.addWidget(self.printer_combo, 0, 1)
        
        self.refresh_btn = QtWidgets.QPushButton("Обновить список")
        self.refresh_btn.clicked.connect(self.find_printers)
        printer_layout.addWidget(self.refresh_btn, 0, 2)
        
        # Настройки размера
        printer_layout.addWidget(QtWidgets.QLabel("Размер (мм):"), 1, 0)
        self.dm_size_spin = QtWidgets.QSpinBox()
        self.dm_size_spin.setRange(15, 50)
        self.dm_size_spin.setValue(30)
        self.dm_size_spin.valueChanged.connect(self.save_settings)
        printer_layout.addWidget(self.dm_size_spin, 1, 1)
        
        printer_layout.addWidget(QtWidgets.QLabel("Тихая зона (мм):"), 2, 0)
        self.quiet_zone_spin = QtWidgets.QSpinBox()
        self.quiet_zone_spin.setRange(1, 10)
        self.quiet_zone_spin.setValue(3)
        self.quiet_zone_spin.valueChanged.connect(self.save_settings)
        printer_layout.addWidget(self.quiet_zone_spin, 2, 1)
        
        # Кнопка тестовой печати
        self.test_print_btn = QtWidgets.QPushButton("Тест печати")
        self.test_print_btn.clicked.connect(self.test_print)
        printer_layout.addWidget(self.test_print_btn, 3, 0, 1, 3)
        
        printer_group.setLayout(printer_layout)
        
        # Форма ввода данных
        input_group = QtWidgets.QGroupBox("Сканирование DataMatrix")
        input_layout = QtWidgets.QVBoxLayout()
        
        self.text_input = QtWidgets.QLineEdit()
        self.text_input.setPlaceholderText("Поднесите DataMatrix к сканеру...")
        self.text_input.textChanged.connect(self.on_text_changed)
        self.text_input.setFixedHeight(50)
        font = self.text_input.font()
        font.setPointSize(12)
        self.text_input.setFont(font)
        
        input_layout.addWidget(QtWidgets.QLabel("Данные сканирования:"))
        input_layout.addWidget(self.text_input)
        input_group.setLayout(input_layout)
        
        # Статусная строка
        self.status_label = QtWidgets.QLabel("Готов к сканированию")
        self.status_label.setAlignment(QtCore.Qt.AlignCenter)
        self.status_label.setStyleSheet("background-color: #e0e0e0; padding: 10px; border-radius: 5px;")
        
        # Лог событий
        log_group = QtWidgets.QGroupBox("Лог событий")
        log_layout = QtWidgets.QVBoxLayout()
        
        self.log_text = QtWidgets.QTextEdit()
        self.log_text.setReadOnly(True)
        
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        
        # Добавляем все группы в основной layout
        layout.addWidget(printer_group)
        layout.addWidget(input_group)
        layout.addWidget(self.status_label)
        layout.addWidget(log_group)
        
        self.text_input.setFocus()
        self.print_timer = QtCore.QTimer()
        self.print_timer.setSingleShot(True)
        self.print_timer.timeout.connect(self.auto_print_dm_code)
        self.current_scanned_data = ""
        
    def init_database(self):
        self.conn = sqlite3.connect('print_history.db')
        self.cursor = self.conn.cursor()
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS print_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                dm_content TEXT,
                printer_name TEXT,
                print_date TIMESTAMP,
                print_success BOOLEAN
            )
        ''')
        self.conn.commit()
    
    def load_settings(self):
        """Загрузка сохраненных настроек"""
        self.dm_size_spin.setValue(self.settings.value("dm_size", 30, type=int))
        self.quiet_zone_spin.setValue(self.settings.value("quiet_zone", 3, type=int))
    
    def save_settings(self):
        """Сохранение настроек"""
        self.settings.setValue("dm_size", self.dm_size_spin.value())
        self.settings.setValue("quiet_zone", self.quiet_zone_spin.value())
    
    def on_printer_changed(self, index):
        """Обработчик изменения выбора принтера"""
        if index >= 0:
            printer_name = self.printer_combo.itemText(index)
            self.log_text.append(f"Выбран принтер: {printer_name}")
    
    def test_print(self):
        """Тестовая печать"""
        test_data = "010461414111072521pNxU640aZ4Kk"
        self.process_dm_code(test_data, is_test=True)
    
    def on_text_changed(self, text):
        self.print_timer.stop()
        current_data = text.strip()
        if current_data:
            self.current_scanned_data = current_data
            self.print_timer.start(500)
    
    def auto_print_dm_code(self):
        """Автоматическая печать после сканирования"""
        if self.current_scanned_data:
            self.process_dm_code(self.current_scanned_data)
            self.current_scanned_data = ""
            self.text_input.clear()
            self.text_input.setFocus()
    
    def find_printers(self):
        self.printer_combo.clear()
        try:
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
            for printer in printers:
                self.printer_combo.addItem(printer[2])
            
            default_printer = win32print.GetDefaultPrinter()
            index = self.printer_combo.findText(default_printer)
            if index >= 0:
                self.printer_combo.setCurrentIndex(index)
                
            self.log_text.append(f"Найдено принтеров: {len(printers)}")
            
        except Exception as e:
            self.log_text.append(f"Ошибка поиска принтеров: {str(e)}")
    
    def generate_data_matrix(self, data, size_mm=30, quiet_zone=3):
        """Простая генерация DataMatrix"""
        try:
            # Кодируем данные
            encoded = encode(data.encode('utf-8'))
            
            # Создаем изображение из encoded data
            img = Image.frombytes('L', (encoded.width, encoded.height), encoded.pixels)
            
            # Конвертируем в черно-белое
            img = img.convert('1')
            
            # Добавляем тихую зону
            quiet_zone_px = int(quiet_zone * 3.78)  # 3.78 пикселя на мм
            if quiet_zone_px > 0:
                new_width = img.width + quiet_zone_px * 2
                new_height = img.height + quiet_zone_px * 2
                new_img = Image.new('1', (new_width, new_height), 1)
                new_img.paste(img, (quiet_zone_px, quiet_zone_px))
                img = new_img
            
            # Масштабируем до нужного размера
            target_size_px = int(size_mm * 3.78)
            img = img.resize((target_size_px, target_size_px), Image.NEAREST)
            
            return img
            
        except Exception as e:
            self.log_text.append(f"Ошибка генерации DataMatrix: {str(e)}")
            return None
    
    def process_dm_code(self, data, is_test=False):
        """Обработка DataMatrix"""
        if not data:
            return False
        
        if self.printer_combo.count() == 0:
            self.status_label.setText("Ошибка: Принтеры не найдены!")
            return False
        
        printer_name = self.printer_combo.currentText()
        
        try:
            # Генерируем DataMatrix
            dm_size = self.dm_size_spin.value()
            quiet_zone = self.quiet_zone_spin.value()
            dm_image = self.generate_data_matrix(data, dm_size, quiet_zone)
            
            if not dm_image:
                self.status_label.setText("Ошибка генерации кода!")
                return False
            
            # Печатаем
            success = self.print_image(dm_image, printer_name)
            
            # Сохраняем в историю
            self.cursor.execute(
                "INSERT INTO print_history (dm_content, printer_name, print_date, print_success) VALUES (?, ?, ?, ?)",
                (data, printer_name, datetime.now(), success)
            )
            self.conn.commit()
            
            if success:
                status_text = "Тестовая печать успешна!" if is_test else "Успешно напечатано!"
                self.status_label.setText(status_text)
                self.log_text.append(f"{status_text}: {data}")
            else:
                self.status_label.setText("Ошибка печати!")
            
            return success
            
        except Exception as e:
            self.status_label.setText("Ошибка печати!")
            self.log_text.append(f"Ошибка: {str(e)}")
            return False
    
    def print_image(self, image, printer_name):
        """Простая печать изображения"""
        try:
            # Сохраняем временно как BMP
            with tempfile.NamedTemporaryFile(suffix='.bmp', delete=False) as temp_file:
                temp_file_path = temp_file.name

            # Конвертируем в RGB и сохраняем
            rgb_image = image.convert('RGB')
            rgb_image.save(temp_file_path, 'BMP')

            # Печатаем
            hprinter = win32print.OpenPrinter(printer_name)
            try:
                printer_dc = win32ui.CreateDC()
                printer_dc.CreatePrinterDC(printer_name)

                printer_dc.StartDoc("DataMatrix Print")
                printer_dc.StartPage()

                bmp = Image.open(temp_file_path)
                dib = ImageWin.Dib(bmp)

                # Получаем размеры области печати
                printable_width = printer_dc.GetDeviceCaps(win32con.HORZRES)
                printable_height = printer_dc.GetDeviceCaps(win32con.VERTRES)

                # Центрируем изображение
                x = (printable_width - bmp.width) // 2
                y = (printable_height - bmp.height) // 2

                # Печатаем
                dib.draw(printer_dc.GetHandleOutput(),
                        (x, y, x + bmp.width, y + bmp.height))

                printer_dc.EndPage()
                printer_dc.EndDoc()
                bmp.close()

                return True
            finally:
                win32print.ClosePrinter(hprinter)
        except Exception as e:
            self.log_text.append(f"Ошибка печати: {str(e)}")
            return False
        finally:
            # Удаляем временный файл
            if os.path.exists(temp_file_path):
                try:
                    os.unlink(temp_file_path)
                except:
                    pass

    def closeEvent(self, event):
        self.conn.close()
        self.save_settings()
        event.accept()

def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    
    window = PrinterApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()