import sys
import sqlite3
import tempfile
import os
import win32print
import win32ui
import win32con
from datetime import datetime
from PIL import Image, ImageWin, ImageDraw
from PyQt5 import QtWidgets, QtCore, QtGui

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
        printer_group = QtWidgets.QGroupBox("Настройки для Честного знака (GS1)")
        printer_layout = QtWidgets.QGridLayout()
        
        printer_layout.addWidget(QtWidgets.QLabel("Принтер:"), 0, 0)
        self.printer_combo = QtWidgets.QComboBox()
        self.printer_combo.currentIndexChanged.connect(self.on_printer_changed)
        printer_layout.addWidget(self.printer_combo, 0, 1)
        
        self.refresh_btn = QtWidgets.QPushButton("Обновить список")
        self.refresh_btn.clicked.connect(self.find_printers)
        printer_layout.addWidget(self.refresh_btn, 0, 2)
        
        # Настройки для Честного знака
        printer_layout.addWidget(QtWidgets.QLabel("Размер (мм):"), 1, 0)
        self.dm_size_spin = QtWidgets.QSpinBox()
        self.dm_size_spin.setRange(15, 50)
        self.dm_size_spin.setValue(25)
        self.dm_size_spin.valueChanged.connect(self.save_settings)
        printer_layout.addWidget(self.dm_size_spin, 1, 1)
        
        printer_layout.addWidget(QtWidgets.QLabel("Тихая зона (мм):"), 2, 0)
        self.quiet_zone_spin = QtWidgets.QSpinBox()
        self.quiet_zone_spin.setRange(1, 10)
        self.quiet_zone_spin.setValue(3)
        self.quiet_zone_spin.valueChanged.connect(self.save_settings)
        printer_layout.addWidget(self.quiet_zone_spin, 2, 1)
        
        # Добавляем FNC1 в начало
        self.add_fnc1_cb = QtWidgets.QCheckBox("Добавить FNC1 (GS1)")
        self.add_fnc1_cb.setChecked(True)
        self.add_fnc1_cb.stateChanged.connect(self.save_settings)
        printer_layout.addWidget(self.add_fnc1_cb, 3, 0, 1, 2)
        
        # Автодобавление разделителей
        self.auto_gs_cb = QtWidgets.QCheckBox("Авто-добавление GS разделителей")
        self.auto_gs_cb.setChecked(True)
        self.auto_gs_cb.stateChanged.connect(self.save_settings)
        printer_layout.addWidget(self.auto_gs_cb, 4, 0, 1, 2)
        
        # Кнопка тестовой печати
        self.test_print_btn = QtWidgets.QPushButton("Тест GS1 DataMatrix")
        self.test_print_btn.clicked.connect(self.test_print_gs1)
        printer_layout.addWidget(self.test_print_btn, 5, 0, 1, 3)
        
        # Информация о формате
        info_label = QtWidgets.QLabel(
            "Для Честного знака используйте GS1 DataMatrix с FNC1 в начале и GS разделителями"
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: blue; font-style: italic;")
        printer_layout.addWidget(info_label, 6, 0, 1, 3)
        
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
                dm_content TEXT UNIQUE,
                printer_name TEXT,
                print_date TIMESTAMP,
                print_success BOOLEAN
            )
        ''')
        self.conn.commit()
    
    def load_settings(self):
        """Загрузка сохраненных настроек"""
        self.dm_size_spin.setValue(self.settings.value("dm_size", 30, type=int))
        self.quiet_zone_spin.setValue(self.settings.value("quiet_zone", 2, type=int))
        self.add_fnc1_cb.setChecked(self.settings.value("add_fnc1", False, type=bool))  # Отключаем FNC1
        self.auto_gs_cb.setChecked(self.settings.value("auto_gs", False, type=bool))   # Отключаем GS разделители
    
    def save_settings(self):
        """Сохранение настроек"""
        self.settings.setValue("dm_size", self.dm_size_spin.value())
        self.settings.setValue("quiet_zone", self.quiet_zone_spin.value())
        self.settings.setValue("add_fnc1", self.add_fnc1_cb.isChecked())
        self.settings.setValue("auto_gs", self.auto_gs_cb.isChecked())
    
    def on_printer_changed(self, index):
        """Обработчик изменения выбора принтера"""
        if index >= 0:
            printer_name = self.printer_combo.itemText(index)
            self.log_text.append(f"Выбран принтер: {printer_name}")
    
    def test_print_gs1(self):
        """Тестовая печать для GS1 DataMatrix"""
        # Тестовые данные в формате Честного знака
        test_data = "010461414111072521pNxU640aZ4Kk"  # Пример данных
        self.process_dm_code(test_data, is_test=True)
    
    def on_text_changed(self, text):
        self.print_timer.stop()
        current_data = text.strip()
        if current_data:
            self.current_scanned_data = current_data
            self.log_text.append(f"[DEBUG] Данные получены: {current_data}")  # Лог
            self.print_timer.start(200)
            self.log_text.append("[DEBUG] Таймер печати запущен")  # Лог

    
    def auto_print_dm_code(self):
        """Автоматическая печать после сканирования"""
        self.log_text.append("[DEBUG] Таймер сработал, запуск auto_print_dm_code")
        if self.current_scanned_data:
            self.log_text.append(f"[DEBUG] Печатаются данные: {self.current_scanned_data}")
            self.log_text.append(f"[DEBUG] Настройки: FNC1={self.add_fnc1_cb.isChecked()}, GS={self.auto_gs_cb.isChecked()}")  # Добавляем логирование настроек
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
    
    def format_gs1_data(self, data):
        """Правильное форматирование данных для GS1 DataMatrix"""
        try:
            # Для Честного знака используем raw данные
            # FNC1 будет обработан правильно при кодировании
            return data.encode('utf-8')
        except Exception as e:
            self.log_text.append(f"Ошибка форматирования: {str(e)}")
            return data.encode('utf-8')
        
    def generate_gs1_data_matrix(self, data, size_mm=25, quiet_zone=3):
        """Упрощенная генерация DataMatrix без сложных преобразований"""
        try:
            # Просто кодируем данные как есть
            encoded = encode(data.encode('utf-8'))
            
            # Создаем изображение напрямую из encoded data
            img = Image.new('L', (encoded.width, encoded.height), 255)  # Белый фон
            
            # Копируем пиксели
            pixels = encoded.pixels
            for i in range(0, len(pixels), 3):
                x = (i // 3) % encoded.width
                y = (i // 3) // encoded.width
                brightness = pixels[i]  # Берем только R канал
                img.putpixel((x, y), brightness)
            
            # Конвертируем в чисто черно-белое
            img = img.convert('1')
            
            # Добавляем тихую зону
            quiet_zone_px = int(quiet_zone * 3.78)
            if quiet_zone_px > 0:
                new_width = img.width + quiet_zone_px * 2
                new_height = img.height + quiet_zone_px * 2
                new_img = Image.new('1', (new_width, new_height), 1)  # Белый фон
                new_img.paste(img, (quiet_zone_px, quiet_zone_px))
                img = new_img
            
            # Масштабируем
            target_size_px = int(size_mm * 3.78)
            img = img.resize((target_size_px, target_size_px), Image.Resampling.NEAREST)
            
            return img
            
        except Exception as e:
            self.log_text.append(f"Ошибка генерации DataMatrix: {str(e)}")
            return None
    
    def process_dm_code(self, data, is_test=False):
        """Обработка DataMatrix - исправленная версия"""
        if not data:
            return False
        
        if self.printer_combo.count() == 0:
            self.status_label.setText("Ошибка: Принтеры не найдены!")
            return False
        
        printer_name = self.printer_combo.currentText()
        
        # Временное отключение GS разделителей для автоматической печати
        original_auto_gs = self.auto_gs_cb.isChecked()
        if not is_test:  # Для автоматической печати отключаем GS разделители
            self.auto_gs_cb.setChecked(False)
        
        try:
            # Генерируем DataMatrix
            dm_size = self.dm_size_spin.value()
            quiet_zone = self.quiet_zone_spin.value()
            dm_image = self.generate_gs1_data_matrix(data, dm_size, quiet_zone)
            
            if not dm_image:
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
                self.log_text.append(f"{status_text}: {data[:20]}...")
            else:
                self.status_label.setText("Ошибка печати!")
            
            return success
            
        except Exception as e:
            self.status_label.setText("Ошибка печати!")
            self.log_text.append(f"Ошибка: {str(e)}")
            return False
        finally:
            # Восстанавливаем оригинальные настройки
            if not is_test:
                self.auto_gs_cb.setChecked(original_auto_gs)
    
    def print_image(self, image, printer_name):
        """Упрощенная печать"""
        try:
            with tempfile.NamedTemporaryFile(suffix='.bmp', delete=False) as temp_file:
                temp_file_path = temp_file.name

            # ✅ Конвертируем в RGB перед сохранением
            rgb_image = image.convert('RGB')
            rgb_image.save(temp_file_path, 'BMP')

            hprinter = win32print.OpenPrinter(printer_name)
            try:
                printer_dc = win32ui.CreateDC()
                printer_dc.CreatePrinterDC(printer_name)

                printer_dc.StartDoc("DataMatrix Print")
                printer_dc.StartPage()

                bmp = Image.open(temp_file_path)
                dib = ImageWin.Dib(bmp)

                printable_width = printer_dc.GetDeviceCaps(win32con.HORZRES)
                printable_height = printer_dc.GetDeviceCaps(win32con.VERTRES)

                x = (printable_width - bmp.width) // 2
                y = (printable_height - bmp.height) // 2

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