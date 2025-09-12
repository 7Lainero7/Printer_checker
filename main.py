import sys
import sqlite3
import tempfile
import os
import win32print
import win32ui
import win32con
from datetime import datetime
from PIL import Image, ImageWin
from PyQt5 import QtWidgets, QtCore, QtGui

try:
    from pylibdmtx.pylibdmtx import encode
    import pylibdmtx
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
        self.setWindowTitle("DataMatrix Printer Check - Профессиональная версия")
        self.setGeometry(100, 100, 900, 800)
        
        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QtWidgets.QVBoxLayout()
        central_widget.setLayout(layout)
        
        # Форма выбора принтера
        printer_group = QtWidgets.QGroupBox("Настройки принтера")
        printer_layout = QtWidgets.QGridLayout()
        
        printer_layout.addWidget(QtWidgets.QLabel("Принтер:"), 0, 0)
        self.printer_combo = QtWidgets.QComboBox()
        self.printer_combo.currentIndexChanged.connect(self.on_printer_changed)
        printer_layout.addWidget(self.printer_combo, 0, 1)
        
        self.refresh_btn = QtWidgets.QPushButton("Обновить список")
        self.refresh_btn.clicked.connect(self.find_printers)
        printer_layout.addWidget(self.refresh_btn, 0, 2)
        
        # Информация о принтере
        self.printer_info_label = QtWidgets.QLabel("Информация о принтере: не выбрано")
        self.printer_info_label.setWordWrap(True)
        printer_layout.addWidget(self.printer_info_label, 1, 0, 1, 3)
        
        # Настройки размера
        printer_layout.addWidget(QtWidgets.QLabel("Размер DataMatrix (мм):"), 2, 0)
        self.dm_size_spin = QtWidgets.QSpinBox()
        self.dm_size_spin.setRange(5, 100)
        self.dm_size_spin.setValue(20)
        self.dm_size_spin.valueChanged.connect(self.save_settings)
        printer_layout.addWidget(self.dm_size_spin, 2, 1)
        
        # Настройки отступов
        printer_layout.addWidget(QtWidgets.QLabel("Отступ слева (мм):"), 3, 0)
        self.margin_left_spin = QtWidgets.QSpinBox()
        self.margin_left_spin.setRange(0, 100)
        self.margin_left_spin.setValue(5)
        self.margin_left_spin.valueChanged.connect(self.save_settings)
        printer_layout.addWidget(self.margin_left_spin, 3, 1)
        
        printer_layout.addWidget(QtWidgets.QLabel("Отступ сверху (мм):"), 4, 0)
        self.margin_top_spin = QtWidgets.QSpinBox()
        self.margin_top_spin.setRange(0, 100)
        self.margin_top_spin.setValue(5)
        self.margin_top_spin.valueChanged.connect(self.save_settings)
        printer_layout.addWidget(self.margin_top_spin, 4, 1)
        
        # Автоматическая калибровка
        self.auto_calibrate_btn = QtWidgets.QPushButton("Автоматическая калибровка")
        self.auto_calibrate_btn.clicked.connect(self.auto_calibrate)
        printer_layout.addWidget(self.auto_calibrate_btn, 5, 0, 1, 3)
        
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
        
        # Статус сканирования
        self.scan_status = QtWidgets.QLabel("Ожидание сканирования...")
        self.scan_status.setAlignment(QtCore.Qt.AlignCenter)
        
        input_layout.addWidget(QtWidgets.QLabel("Данные сканирования:"))
        input_layout.addWidget(self.text_input)
        input_layout.addWidget(self.scan_status)
        input_group.setLayout(input_layout)
        
        # Статусная строка
        self.status_label = QtWidgets.QLabel("Готов к сканированию")
        self.status_label.setAlignment(QtCore.Qt.AlignCenter)
        self.status_label.setStyleSheet("background-color: #e0e0e0; padding: 10px; border-radius: 5px;")
        
        # Лог событий
        log_group = QtWidgets.QGroupBox("Лог событий и отладка")
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
        
        # Устанавливаем фокус на поле ввода
        self.text_input.setFocus()
        
        # Таймер для автоматической печати
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
        self.dm_size_spin.setValue(self.settings.value("dm_size", 20, type=int))
        self.margin_left_spin.setValue(self.settings.value("margin_left", 5, type=int))
        self.margin_top_spin.setValue(self.settings.value("margin_top", 5, type=int))
    
    def save_settings(self):
        """Сохранение настроек"""
        self.settings.setValue("dm_size", self.dm_size_spin.value())
        self.settings.setValue("margin_left", self.margin_left_spin.value())
        self.settings.setValue("margin_top", self.margin_top_spin.value())
    
    def on_printer_changed(self, index):
        """Обработчик изменения выбора принтера"""
        if index >= 0:
            printer_name = self.printer_combo.itemText(index)
            self.update_printer_info(printer_name)
    
    def update_printer_info(self, printer_name):
        """Обновление информации о выбранном принтере"""
        try:
            hprinter = win32print.OpenPrinter(printer_name)
            printer_info = win32print.GetPrinter(hprinter, 2)
            win32print.ClosePrinter(hprinter)
            
            info_text = f"""
            Принтер: {printer_name}
            Статус: {printer_info['Status']}
            Драйвер: {printer_info['pDriverName']}
            Порт: {printer_info['pPortName']}
            """
            
            self.printer_info_label.setText(info_text.strip())
            self.current_printer_info = printer_info
            
        except Exception as e:
            self.printer_info_label.setText(f"Ошибка получения информации: {str(e)}")
    
    def auto_calibrate(self):
        """Автоматическая калибровка под принтер"""
        if not self.current_printer_info:
            QtWidgets.QMessageBox.warning(self, "Ошибка", "Сначала выберите принтер!")
            return
        
        try:
            # Тестовая печать для калибровки
            test_data = "CALIBRATION_TEST_123"
            self.process_dm_code(test_data, is_test=True, is_calibration=True)
            
        except Exception as e:
            self.log_text.append(f"Ошибка калибровки: {str(e)}")
    
    def on_text_changed(self, text):
        """Обработчик изменения текста для автоматического определения конца сканирования"""
        if text.strip():
            self.current_scanned_data = text.strip()
            self.scan_status.setText("Сканирование...")
            self.scan_status.setStyleSheet("color: blue;")
            # Запускаем таймер - если в течение 200ms не будет изменений, значит сканирование завершено
            self.print_timer.start(200)
    
    def auto_print_dm_code(self):
        """Автоматическая печать после сканирования"""
        if self.current_scanned_data:
            self.scan_status.setText("Обработка...")
            self.scan_status.setStyleSheet("color: orange;")
            self.process_dm_code(self.current_scanned_data)
            self.current_scanned_data = ""
            # Сбрасываем курсор в начало
            self.text_input.setFocus()
            self.text_input.clear()
            self.scan_status.setText("Ожидание сканирования...")
            self.scan_status.setStyleSheet("color: black;")
    
    def find_printers(self):
        self.printer_combo.clear()
        try:
            printers = win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS, 
                None, 1
            )
            for printer in printers:
                self.printer_combo.addItem(printer[2])
            
            # Устанавливаем принтер по умолчанию
            default_printer = win32print.GetDefaultPrinter()
            index = self.printer_combo.findText(default_printer)
            if index >= 0:
                self.printer_combo.setCurrentIndex(index)
                self.update_printer_info(default_printer)
                
            self.log_text.append(f"Найдено принтеров: {len(printers)}")
            
        except Exception as e:
            self.log_text.append(f"Ошибка поиска принтеров: {str(e)}")
    
    def generate_data_matrix(self, data, size_mm=20):
        """Генерация DataMatrix кода"""
        try:
            # Генерируем DataMatrix
            encoded = encode(data.encode('utf-8'))
            
            # Конвертируем в PIL Image
            img = Image.frombytes('RGB', (encoded.width, encoded.height), encoded.pixels)
            
            # Конвертируем в черно-белое
            img = img.convert('L')
            
            # Инвертируем цвета (DataMatrix обычно черный на белом)
            img = img.point(lambda x: 0 if x > 128 else 255)  # Четкое черно-белое
            
            # Масштабируем до нужного размера в мм
            target_size_px = int(size_mm * 3.78)  # 96 DPI
            img = img.resize((target_size_px, target_size_px), Image.Resampling.LANCZOS)
            
            return img
            
        except Exception as e:
            self.log_text.append(f"Ошибка генерации DataMatrix: {str(e)}")
            return None
    
    def process_dm_code(self, data, is_test=False, is_calibration=False):
        """Обработка DataMatrix с проверкой дубликатов"""
        if not data:
            return False
        
        if self.printer_combo.count() == 0:
            self.status_label.setText("Ошибка: Принтеры не найдены!")
            self.status_label.setStyleSheet("background-color: #ffcccc; padding: 10px; border-radius: 5px;")
            return False
        
        printer_name = self.printer_combo.currentText()
        
        # Для тестовой печати пропускаем проверку дубликатов
        if not is_test and not is_calibration:
            # Проверка на дубликат
            try:
                self.cursor.execute(
                    "SELECT * FROM print_history WHERE dm_content = ? AND print_success = 1",
                    (data,)
                )
                duplicate = self.cursor.fetchone()
                
                if duplicate:
                    self.status_label.setText("Дубликат! Печать отменена")
                    self.status_label.setStyleSheet("background-color: #ffffcc; padding: 10px; border-radius: 5px;")
                    self.log_text.append(f"Дубликат обнаружен: {data[:30]}...")
                    self.text_input.clear()
                    self.text_input.setFocus()
                    return False
                    
            except Exception as e:
                self.log_text.append(f"Ошибка проверки дубликата: {str(e)}")
                return False
        
        try:
            # Генерируем DataMatrix с настройками размера
            dm_size = self.dm_size_spin.value()
            dm_image = self.generate_data_matrix(data, dm_size)
            if not dm_image:
                return False
            
            # Печатаем DataMatrix с настройками отступов
            margin_left = self.margin_left_spin.value()
            margin_top = self.margin_top_spin.value()
            
            success = self.print_image(dm_image, printer_name, margin_left, margin_top)
            
            # Сохраняем в историю
            self.cursor.execute(
                "INSERT INTO print_history (dm_content, printer_name, print_date, print_success) VALUES (?, ?, ?, ?)",
                (data, printer_name, datetime.now(), success)
            )
            self.conn.commit()
            
            if success:
                if is_calibration:
                    status_text = "Калибровка успешна!"
                elif is_test:
                    status_text = "Тестовая печать успешна!"
                else:
                    status_text = "Успешно напечатано!"
                
                self.status_label.setText(status_text)
                self.status_label.setStyleSheet("background-color: #ccffcc; padding: 10px; border-radius: 5px;")
                self.log_text.append(f"{status_text}: {data[:30]}...")
            else:
                self.status_label.setText("Ошибка печати!")
                self.status_label.setStyleSheet("background-color: #ffcccc; padding: 10px; border-radius: 5px;")
            
            # Сбрасываем фокус и очищаем поле
            self.text_input.clear()
            self.text_input.setFocus()
            return success
            
        except Exception as e:
            self.status_label.setText("Ошибка печати!")
            self.status_label.setStyleSheet("background-color: #ffcccc; padding: 10px; border-radius: 5px;")
            self.log_text.append(f"Ошибка печати: {str(e)}")
            
            # Сохраняем ошибку в базу
            try:
                self.cursor.execute(
                    "INSERT INTO print_history (dm_content, printer_name, print_date, print_success) VALUES (?, ?, ?, ?)",
                    (data, printer_name, datetime.now(), False)
                )
                self.conn.commit()
            except:
                pass
                
            return False
    
    def print_image(self, image, printer_name, margin_left_mm=5, margin_top_mm=5):
        """Печать изображения с настройкой отступов"""
        temp_file_path = None
        try:
            # Создаем временный файл
            temp_file = tempfile.NamedTemporaryFile(suffix='.bmp', delete=False)
            temp_file_path = temp_file.name
            temp_file.close()
            
            # Сохраняем изображение
            image.save(temp_file_path, 'BMP')
            
            # Создаем контекст устройства принтера
            printer_dc = win32ui.CreateDC()
            printer_dc.CreatePrinterDC(printer_name)
            
            # Получаем информацию о принтере
            dpi_x = printer_dc.GetDeviceCaps(win32con.LOGPIXELSX)
            dpi_y = printer_dc.GetDeviceCaps(win32con.LOGPIXELSY)
            printable_width = printer_dc.GetDeviceCaps(win32con.HORZRES)
            printable_height = printer_dc.GetDeviceCaps(win32con.VERTRES)
            
            self.log_text.append(f"DPI принтера: {dpi_x}x{dpi_y}")
            self.log_text.append(f"Область печати: {printable_width}x{printable_height} пикселей")
            
            printer_dc.StartDoc("DataMatrix Print")
            printer_dc.StartPage()
            
            # Конвертируем мм в пиксели
            pixels_per_mm_x = dpi_x / 25.4
            pixels_per_mm_y = dpi_y / 25.4
            
            margin_left_px = int(margin_left_mm * pixels_per_mm_x)
            margin_top_px = int(margin_top_mm * pixels_per_mm_y)
            
            # Загружаем изображение
            bmp = Image.open(temp_file_path)
            
            # Проверяем, помещается ли изображение
            if margin_left_px + bmp.width > printable_width:
                self.log_text.append("Предупреждение: Изображение выходит за правую границу!")
            if margin_top_px + bmp.height > printable_height:
                self.log_text.append("Предупреждение: Изображение выходит за нижнюю границу!")
            
            # Рисуем изображение с отступами
            dib = ImageWin.Dib(bmp)
            dib.draw(printer_dc.GetHandleOutput(), 
                    (margin_left_px, margin_top_px, 
                     margin_left_px + bmp.width, margin_top_px + bmp.height))
            
            printer_dc.EndPage()
            printer_dc.EndDoc()
            
            # Закрываем изображение
            bmp.close()
            
            self.log_text.append("Печать выполнена успешно")
            return True
                
        except Exception as e:
            self.log_text.append(f"Ошибка печати: {str(e)}")
            return False
        finally:
            # Удаляем временный файл
            if temp_file_path and os.path.exists(temp_file_path):
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