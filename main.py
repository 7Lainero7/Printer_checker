import sys
import sqlite3
import qrcode
import tempfile
import os
import win32print
import win32ui
import win32con
from datetime import datetime
from PIL import Image, ImageWin
from PyQt5 import QtWidgets, QtCore, QtGui

class PrinterApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QtCore.QSettings("PrinterCheck", "QRPrinter")
        self.init_ui()
        self.init_database()
        self.find_printers()
        self.load_settings()
        
    def init_ui(self):
        self.setWindowTitle("QR Printer Check")
        self.setGeometry(100, 100, 800, 700)
        
        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QtWidgets.QVBoxLayout()
        central_widget.setLayout(layout)
        
        # Форма выбора принтера
        printer_group = QtWidgets.QGroupBox("Настройки принтера")
        printer_layout = QtWidgets.QGridLayout()
        
        printer_layout.addWidget(QtWidgets.QLabel("Принтер:"), 0, 0)
        self.printer_combo = QtWidgets.QComboBox()
        printer_layout.addWidget(self.printer_combo, 0, 1)
        
        self.refresh_btn = QtWidgets.QPushButton("Обновить список")
        self.refresh_btn.clicked.connect(self.find_printers)
        printer_layout.addWidget(self.refresh_btn, 0, 2)
        
        # Настройки размера
        printer_layout.addWidget(QtWidgets.QLabel("Размер QR-кода (мм):"), 1, 0)
        self.qr_size_spin = QtWidgets.QSpinBox()
        self.qr_size_spin.setRange(10, 100)
        self.qr_size_spin.setValue(50)
        self.qr_size_spin.valueChanged.connect(self.save_settings)
        printer_layout.addWidget(self.qr_size_spin, 1, 1)
        
        # Настройки отступов
        printer_layout.addWidget(QtWidgets.QLabel("Отступ слева (мм):"), 2, 0)
        self.margin_left_spin = QtWidgets.QSpinBox()
        self.margin_left_spin.setRange(0, 100)
        self.margin_left_spin.setValue(10)
        self.margin_left_spin.valueChanged.connect(self.save_settings)
        printer_layout.addWidget(self.margin_left_spin, 2, 1)
        
        printer_layout.addWidget(QtWidgets.QLabel("Отступ сверху (мм):"), 3, 0)
        self.margin_top_spin = QtWidgets.QSpinBox()
        self.margin_top_spin.setRange(0, 100)
        self.margin_top_spin.setValue(10)
        self.margin_top_spin.valueChanged.connect(self.save_settings)
        printer_layout.addWidget(self.margin_top_spin, 3, 1)
        
        # Кнопка тестовой печати
        self.test_print_btn = QtWidgets.QPushButton("Тестовая печать")
        self.test_print_btn.clicked.connect(self.test_print)
        printer_layout.addWidget(self.test_print_btn, 4, 0, 1, 3)
        
        printer_group.setLayout(printer_layout)
        
        # Форма ввода данных
        input_group = QtWidgets.QGroupBox("Сканирование QR-кодов")
        input_layout = QtWidgets.QVBoxLayout()
        
        self.text_input = QtWidgets.QLineEdit()
        self.text_input.setPlaceholderText("Поднесите QR-код к сканеру...")
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
        
        # Устанавливаем фокус на поле ввода
        self.text_input.setFocus()
        
        # Таймер для автоматической печати
        self.print_timer = QtCore.QTimer()
        self.print_timer.setSingleShot(True)
        self.print_timer.timeout.connect(self.auto_print_qr_code)
        
        self.current_scanned_data = ""
        
    def init_database(self):
        self.conn = sqlite3.connect('print_history.db')
        self.cursor = self.conn.cursor()
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS print_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                qr_content TEXT UNIQUE,
                printer_name TEXT,
                print_date TIMESTAMP
            )
        ''')
        self.conn.commit()
    
    def load_settings(self):
        """Загрузка сохраненных настроек"""
        self.qr_size_spin.setValue(self.settings.value("qr_size", 50, type=int))
        self.margin_left_spin.setValue(self.settings.value("margin_left", 10, type=int))
        self.margin_top_spin.setValue(self.settings.value("margin_top", 10, type=int))
    
    def save_settings(self):
        """Сохранение настроек"""
        self.settings.setValue("qr_size", self.qr_size_spin.value())
        self.settings.setValue("margin_left", self.margin_left_spin.value())
        self.settings.setValue("margin_top", self.margin_top_spin.value())
    
    def on_text_changed(self, text):
        """Обработчик изменения текста для автоматического определения конца сканирования"""
        if text.strip():
            self.current_scanned_data = text.strip()
            # Запускаем таймер - если в течение 500ms не будет изменений, значит сканирование завершено
            self.print_timer.start(500)
    
    def auto_print_qr_code(self):
        """Автоматическая печать после сканирования"""
        if self.current_scanned_data:
            self.process_qr_code(self.current_scanned_data)
            self.current_scanned_data = ""
            # Сбрасываем курсор в начало
            self.text_input.setFocus()
            self.text_input.clear()
    
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
                
            self.log_text.append(f"Найдено принтеров: {len(printers)}")
            
        except Exception as e:
            self.log_text.append(f"Ошибка поиска принтеров: {str(e)}")
    
    def test_print(self):
        """Тестовая печать для настройки"""
        test_data = "TEST_QR_CODE_1234567890"
        self.process_qr_code(test_data, is_test=True)
    
    def generate_qr_code(self, data, size_mm=50):
        """Генерация QR-кода с указанием размера в мм"""
        try:
            # Конвертируем мм в пиксели (примерно 3.78 пикселя на мм для 96 DPI)
            size_pixels = int(size_mm * 3.78)
            
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_H,  # Высокий уровень коррекции
                box_size=10,
                border=4,
            )
            qr.add_data(data)
            qr.make(fit=True)
            
            img = qr.make_image(fill_color="black", back_color="white")
            # Масштабируем до нужного размера
            img = img.resize((size_pixels, size_pixels), Image.Resampling.LANCZOS)
            return img
        except Exception as e:
            self.log_text.append(f"Ошибка генерации QR-кода: {str(e)}")
            return None
    
    def process_qr_code(self, data, is_test=False):
        """Обработка QR-кода с проверкой дубликатов"""
        if not data:
            return False
        
        if self.printer_combo.count() == 0:
            self.status_label.setText("Ошибка: Принтеры не найдены!")
            self.status_label.setStyleSheet("background-color: #ffcccc; padding: 10px; border-radius: 5px;")
            return False
        
        printer_name = self.printer_combo.currentText()
        
        # Для тестовой печати пропускаем проверку дубликатов
        if not is_test:
            # Проверка на дубликат
            try:
                self.cursor.execute(
                    "SELECT * FROM print_history WHERE qr_content = ?",
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
            # Генерируем QR-код с настройками размера
            qr_size = self.qr_size_spin.value()
            qr_image = self.generate_qr_code(data, qr_size)
            if not qr_image:
                return False
            
            # Печатаем QR-код с настройками отступов
            margin_left = self.margin_left_spin.value()
            margin_top = self.margin_top_spin.value()
            self.print_image(qr_image, printer_name, margin_left, margin_top)
            
            # Сохраняем в историю (кроме тестовой печати)
            if not is_test:
                self.cursor.execute(
                    "INSERT INTO print_history (qr_content, printer_name, print_date) VALUES (?, ?, ?)",
                    (data, printer_name, datetime.now())
                )
                self.conn.commit()
            
            status_text = "Тестовая печать успешна!" if is_test else "Успешно напечатано!"
            self.status_label.setText(status_text)
            self.status_label.setStyleSheet("background-color: #ccffcc; padding: 10px; border-radius: 5px;")
            self.log_text.append(f"{status_text}: {data[:30]}...")
            
            # Сбрасываем фокус и очищаем поле
            self.text_input.clear()
            self.text_input.setFocus()
            return True
            
        except Exception as e:
            self.status_label.setText("Ошибка печати!")
            self.status_label.setStyleSheet("background-color: #ffcccc; padding: 10px; border-radius: 5px;")
            self.log_text.append(f"Ошибка печати: {str(e)}")
            return False
    
    def print_image(self, image, printer_name, margin_left_mm=10, margin_top_mm=10):
        """Печать изображения с настройкой отступов"""
        try:
            # Конвертируем PIL Image в формат для печати
            temp_file = tempfile.NamedTemporaryFile(suffix='.bmp', delete=False)
            image.save(temp_file.name, 'BMP')
            temp_file.close()
            
            # Получаем настройки принтера
            hprinter = win32print.OpenPrinter(printer_name)
            printer_info = win32print.GetPrinter(hprinter, 2)
            win32print.ClosePrinter(hprinter)
            
            # Создаем контекст устройства принтера
            printer_dc = win32ui.CreateDC()
            printer_dc.CreatePrinterDC(printer_name)
            printer_dc.StartDoc("QR Code Print")
            printer_dc.StartPage()
            
            # Конвертируем мм в пиксели (96 DPI = 3.78 пикс/мм)
            dpi = printer_dc.GetDeviceCaps(88)  # LOGPIXELSX
            pixels_per_mm = dpi / 25.4
            
            margin_left_px = int(margin_left_mm * pixels_per_mm)
            margin_top_px = int(margin_top_mm * pixels_per_mm)
            
            # Загружаем изображение
            bmp = Image.open(temp_file.name)
            
            # Масштабируем изображение
            width = bmp.width
            height = bmp.height
            
            # Рисуем изображение с отступами
            dib = ImageWin.Dib(bmp)
            dib.draw(printer_dc.GetHandleOutput(), 
                    (margin_left_px, margin_top_px, 
                     margin_left_px + width, margin_top_px + height))
            
            printer_dc.EndPage()
            printer_dc.EndDoc()
            
            os.unlink(temp_file.name)
                
        except Exception as e:
            raise Exception(f"Ошибка печати изображения: {str(e)}")
    
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