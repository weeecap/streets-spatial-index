import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QLineEdit, QPushButton, QGroupBox, QProgressBar,
                            QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem,
                            QHeaderView, QTextEdit, QTabWidget)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont

from nomenclatural import NomenclaturalStreetIndexer

class ProcessingThread(QThread):

    progress_updated = pyqtSignal(int)          
    finished_processing = pyqtSignal(pd.DataFrame)  
    error_occurred = pyqtSignal(str)            
    
    def __init__(self, indexer, file_path, x_col, y_col, street_col=None):
        super().__init__()
        self.indexer = indexer
        self.file_path = file_path
        self.x_col = x_col
        self.y_col = y_col
        self.street_col = street_col  
        self.df_result = None
        
    def run(self):

        try:
            original_df = pd.read_excel(self.file_path)
            
            if self.x_col not in original_df.columns or self.y_col not in original_df.columns:
                self.error_occurred.emit(f"Столбцы '{self.x_col}' и/или '{self.y_col}' не найдены в файле")
                return
            
            if self.street_col and self.street_col not in original_df.columns:
                self.error_occurred.emit(f"Столбец '{self.street_col}' не найден в файле")
                return
            
            nomenclatural_indices = []
            list_numbers = []
            formatted_streets = []
            total_rows = len(original_df)
            
            for idx, row in original_df.iterrows():
                try:
                    x_val = float(row[self.x_col])
                    y_val = float(row[self.y_col])
                    
                    nomenclatural_index = self.indexer.calculate_nomenclatural_index(x_val, y_val)
                    list_number = self.indexer.calculate_list_number(x_val, y_val)
                    
                    if self.street_col:
                        street_name = str(row[self.street_col])
                        formatted_street = self.indexer.format_street_name(street_name)
                    else:
                        formatted_street = ""
                    
                    nomenclatural_indices.append(nomenclatural_index)
                    list_numbers.append(list_number)
                    formatted_streets.append(formatted_street)
                    
                except (ValueError, TypeError):
                    nomenclatural_indices.append("Ошибка")
                    list_numbers.append("Ошибка")
                    formatted_streets.append("Ошибка")
                
                progress = int((idx + 1) / total_rows * 100)
                self.progress_updated.emit(progress)
            
            result_df = pd.DataFrame()
            
            result_df['Номенклатурный_индекс'] = nomenclatural_indices
            result_df['Лист_карты'] = list_numbers
            result_df['Форматированная_улица'] = formatted_streets
            
            self.df_result = result_df
            self.finished_processing.emit(result_df)
            
        except Exception as e:
            self.error_occurred.emit(str(e))
                    
class ExcelProcessorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.indexer = NomenclaturalStreetIndexer(500)
        self.file_path = None
        self.current_df = None
        self.processing_thread = None
        
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Обработчик номенклатурных индексов")
        self.setGeometry(100, 100, 800, 700)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout(central_widget)
        
        title_label = QLabel("Обработка номенклатурных индексов")
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        tabs = QTabWidget()
        layout.addWidget(tabs)
        
        processing_tab = QWidget()
        tabs.addTab(processing_tab, "Обработка")
        
        self.preview_tab = QWidget()
        tabs.addTab(self.preview_tab, "Предпросмотр")
        
        self.setup_processing_tab(processing_tab)
        self.setup_preview_tab(self.preview_tab)
        
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Готов к работе")
        
    def setup_processing_tab(self, tab):
        layout = QVBoxLayout(tab)
        
        origin_group = self.create_origin_group()
        layout.addWidget(origin_group)
        
        file_group = self.create_file_group()
        layout.addWidget(file_group)
        
        settings_group = self.create_settings_group()
        layout.addWidget(settings_group)
        
        self.process_btn = QPushButton("Обработать файл")
        self.process_btn.setStyleSheet(self.get_process_button_style())
        self.process_btn.clicked.connect(self.process_file)
        self.process_btn.setEnabled(False)
        layout.addWidget(self.process_btn)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        instruction_group = self.create_instruction_group()
        layout.addWidget(instruction_group)
        
    def create_origin_group(self):
        """Создает группу элементов для установки начала координат."""
        group = QGroupBox("Начало координат")
        layout = QHBoxLayout(group)
        
        layout.addWidget(QLabel("X:"))
        self.x_entry = QLineEdit()
        self.x_entry.setPlaceholderText("Введите координату X")
        layout.addWidget(self.x_entry)
        
        layout.addWidget(QLabel("Y:"))
        self.y_entry = QLineEdit()
        self.y_entry.setPlaceholderText("Введите координату Y")
        layout.addWidget(self.y_entry)
        
        self.set_origin_btn = QPushButton("Установить начало координат")
        self.set_origin_btn.clicked.connect(self.set_origin)
        layout.addWidget(self.set_origin_btn)
        
        return group
    
    def create_file_group(self):
        """Создает группу элементов для работы с файлами."""
        group = QGroupBox("Работа с файлом")
        layout = QVBoxLayout(group)
        
        self.load_file_btn = QPushButton("Загрузить Excel файл")
        self.load_file_btn.clicked.connect(self.load_file)
        layout.addWidget(self.load_file_btn)
        
        self.file_label = QLabel("Файл не выбран")
        self.file_label.setStyleSheet("color: red;")
        layout.addWidget(self.file_label)
        
        return group
    
    def create_settings_group(self):
        group = QGroupBox("Настройки обработки")
        layout = QHBoxLayout(group)
        
        layout.addWidget(QLabel("Столбец X:"))
        self.x_col_entry = QLineEdit()
        self.x_col_entry.setText("X")
        layout.addWidget(self.x_col_entry)
        
        layout.addWidget(QLabel("Столбец Y:"))
        self.y_col_entry = QLineEdit()
        self.y_col_entry.setText("Y")
        layout.addWidget(self.y_col_entry)

        layout.addWidget(QLabel('Улицы:'))
        self.street_entry = QLineEdit()
        self.street_entry.setText('SEM9')  
        layout.addWidget(self.street_entry)
        
        return group
    
    def create_instruction_group(self):
        group = QGroupBox("Инструкция")
        layout = QVBoxLayout(group)
        
        instruction_text = QTextEdit()
        instruction_text.setPlainText("""1. Установите начало координат (X и Y)
2. Загрузите Excel файл с координатами
3. Укажите названия столбцов с координатами (по умолчанию X и Y)
4. Укажите название столбца с улицами (опционально)
5. Нажмите 'Обработать файл'
6. Выберите место для сохранения результата""")
        instruction_text.setReadOnly(True)
        layout.addWidget(instruction_text)
        
        return group
    
    def setup_preview_tab(self, tab):
        layout = QVBoxLayout(tab)
        
        self.preview_label = QLabel("Данные для предпросмотра появятся после обработки файла")
        self.preview_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.preview_label)
        
        self.preview_table = QTableWidget()
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.preview_table.setVisible(False)
        layout.addWidget(self.preview_table)
    
    def get_process_button_style(self):
        return """
            QPushButton {
                background-color: lightblue;
                font-weight: bold;
                font-size: 14px;
                padding: 10px;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """
    
    def set_origin(self):
        try:
            x = float(self.x_entry.text())
            y = float(self.y_entry.text())
            self.indexer.set_origin(x, y)
            QMessageBox.information(self, "Успех", f"Начало координат установлено: X={x}, Y={y}")
            self.status_bar.showMessage("Начало координат установлено")
        except ValueError:
            QMessageBox.critical(self, "Ошибка", "Введите корректные числовые значения для координат")
    
    def load_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите Excel файл",
            "",
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        
        if file_path:
            self.file_path = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.file_label.setStyleSheet("color: green;")
            self.process_btn.setEnabled(True)
            self.status_bar.showMessage("Файл загружен, готов к обработке")
    
    def process_file(self):
        if not self.file_path:
            QMessageBox.critical(self, "Ошибка", "Сначала загрузите файл")
            return
            
        if self.indexer.origin_x is None or self.indexer.origin_y is None:
            QMessageBox.critical(self, "Ошибка", "Сначала установите начало координат")
            return
        
        x_col = self.x_col_entry.text().strip()
        y_col = self.y_col_entry.text().strip()
        street_col = self.street_entry.text().strip() 
        
        self.processing_thread = ProcessingThread(self.indexer, self.file_path, x_col, y_col, street_col)
        self.processing_thread.progress_updated.connect(self.update_progress)
        self.processing_thread.finished_processing.connect(self.on_processing_finished)
        self.processing_thread.error_occurred.connect(self.on_processing_error)
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.process_btn.setEnabled(False)
        self.status_bar.showMessage("Обработка файла...")
        
        self.processing_thread.start()
    
    def update_progress(self, value):
        self.progress_bar.setValue(value)
    
    def on_processing_finished(self, df):
        self.current_df = df
        self.progress_bar.setVisible(False)
        self.process_btn.setEnabled(True)
        self.status_bar.showMessage("Обработка завершена")
        
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить результат как",
            "",
            "Excel files (*.xlsx);;All files (*.*)"
        )
        
        if output_path:
            try:
                if not output_path.endswith('.xlsx'):
                    output_path += '.xlsx'
                
                df.to_excel(output_path, index=False)
                QMessageBox.information(self, "Успех", f"Файл сохранен как:\n{output_path}")
                self.status_bar.showMessage("Файл успешно обработан и сохранен")
                
                self.show_preview(df)
                
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при сохранении файла: {str(e)}")
        else:
            self.status_bar.showMessage("Обработка отменена")
    
    def on_processing_error(self, error_message):
        self.progress_bar.setVisible(False)
        self.process_btn.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при обработке: {error_message}")
        self.status_bar.showMessage("Ошибка обработки")
    
    def show_preview(self, df):
        tab_widget = self.centralWidget().layout().itemAt(1).widget()
        tab_widget.setCurrentIndex(1)
        
        self.preview_table.setVisible(True)
        self.preview_label.setVisible(False)
        
        preview_rows = min(10, len(df))
        self.preview_table.setRowCount(preview_rows)
        self.preview_table.setColumnCount(len(df.columns))
        self.preview_table.setHorizontalHeaderLabels(df.columns.tolist())
        
        for i in range(preview_rows):
            for j, column in enumerate(df.columns):
                item = QTableWidgetItem(str(df.iloc[i, j]))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.preview_table.setItem(i, j, item)