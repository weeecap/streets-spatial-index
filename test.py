import math
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Tuple 
import os

class NomenclaturalStreetIndexer:
    def __init__(self, square_size: int = 500):
        """
        Args:
            square_size: размер квадрата в метрах (по умолчанию 500м)
        """
        self.square_size = square_size
        self.origin_x = None
        self.origin_y = None
        
        # Алфавит для буквенной части индекса (кириллица)
        self.letters = [chr(i) for i in range(1040, 1072)]  # А-Я
        self.letters = [letter for letter in self.letters if letter != 'Ё' and letter != 'Й']
        
    def set_origin(self, x: float, y: float):
        self.origin_x = x
        self.origin_y = y
        print(f"Начало координат установлено: X={x}, Y={y}")

    def _calculate_(self, x: float, y: float) -> tuple:
        """Вычисляет дельты и индексы для координат"""
        if self.origin_x is None or self.origin_y is None:
            raise ValueError("Начало координат не установлено!")
        
        delta_x = self.origin_x - x 
        delta_y = y - self.origin_y
        
        col_index = math.floor(delta_x / self.square_size)
        row_index = math.floor(delta_y / self.square_size)
        
        return delta_x, delta_y, col_index, row_index
    
    def calculate_nomenclatural_index(self, x: float, y: float) -> str:
        """
        Args:
            x: координата X
            y: координата Y
            
        Returns:
            Номенклатурный индекс (строка)
        """
        delta_x, delta_y, col_index, row_index = self._calculate_(x, y)

        col_index = math.floor(delta_x / self.square_size)
        row_index = math.floor(delta_y / self.square_size)

        letter = self.letters[col_index % len(self.letters)]
        number = row_index + 1  
        
        return f"{letter}-{number}"
    
    def calculate_list_number(self, x: float, y: float, border: int = 4500) -> str:
        """
        Args:
            x: координата X
            y: координата Y
            border: граница для определения листа
            
        Returns: 
            Лист карты 
        """
        delta_x, delta_y, col_index, row_index = self._calculate_(x, y)
        number = row_index + 1
        
        if number <= 9:
            list_number = "Лист 1" if delta_x <= border else "Лист 3"
        else:
            list_number = "Лист 2" if delta_x <= border else "Лист 4"

        return list_number

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработчик номенклатурных индексов")
        self.root.geometry("600x500")
        
        self.indexer = NomenclaturalStreetIndexer(500)
        self.file_path = None
        
        self.create_widgets()
        
    def create_widgets(self):
        # Заголовок
        title_label = tk.Label(self.root, text="Обработка номенклатурных индексов", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Фрейм для ввода начала координат
        origin_frame = tk.LabelFrame(self.root, text="Начало координат", padx=10, pady=10)
        origin_frame.pack(pady=10, padx=20, fill="x")
        
        tk.Label(origin_frame, text="X:").grid(row=0, column=0, padx=5)
        self.x_entry = tk.Entry(origin_frame, width=15)
        self.x_entry.grid(row=0, column=1, padx=5)
        
        tk.Label(origin_frame, text="Y:").grid(row=0, column=2, padx=5)
        self.y_entry = tk.Entry(origin_frame, width=15)
        self.y_entry.grid(row=0, column=3, padx=5)
        
        self.set_origin_btn = tk.Button(origin_frame, text="Установить начало координат", 
                                       command=self.set_origin)
        self.set_origin_btn.grid(row=0, column=4, padx=10)
        
        # Фрейм для работы с файлом
        file_frame = tk.LabelFrame(self.root, text="Работа с файлом", padx=10, pady=10)
        file_frame.pack(pady=10, padx=20, fill="x")
        
        self.load_file_btn = tk.Button(file_frame, text="Загрузить Excel файл", 
                                      command=self.load_file)
        self.load_file_btn.pack(pady=5)
        
        self.file_label = tk.Label(file_frame, text="Файл не выбран", fg="red")
        self.file_label.pack(pady=5)
        
        # Фрейм для настроек обработки
        settings_frame = tk.LabelFrame(self.root, text="Настройки обработки", padx=10, pady=10)
        settings_frame.pack(pady=10, padx=20, fill="x")
        
        tk.Label(settings_frame, text="Столбец X:").grid(row=0, column=0, padx=5)
        self.x_col_entry = tk.Entry(settings_frame, width=10)
        self.x_col_entry.insert(0, "X")
        self.x_col_entry.grid(row=0, column=1, padx=5)
        
        tk.Label(settings_frame, text="Столбец Y:").grid(row=0, column=2, padx=5)
        self.y_col_entry = tk.Entry(settings_frame, width=10)
        self.y_col_entry.insert(0, "Y")
        self.y_col_entry.grid(row=0, column=3, padx=5)
        
        # Кнопка обработки
        self.process_btn = tk.Button(self.root, text="Обработать файл", 
                                    command=self.process_file, state="disabled",
                                    bg="lightblue", font=("Arial", 12, "bold"))
        self.process_btn.pack(pady=20)
        
        # Прогресс бар
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.pack(pady=10, padx=20, fill="x")
        
        # Статус
        self.status_label = tk.Label(self.root, text="Готов к работе", fg="green")
        self.status_label.pack(pady=5)
        
        # Инструкция
        instruction_text = """
Инструкция:
1. Установите начало координат (X и Y)
2. Загрузите Excel файл с координатами
3. Укажите названия столбцов с координатами (по умолчанию X и Y)
4. Нажмите 'Обработать файл'
5. Выберите место для сохранения результата
        """
        instruction_label = tk.Label(self.root, text=instruction_text, justify=tk.LEFT, fg="gray")
        instruction_label.pack(pady=10)
        
    def set_origin(self):
        try:
            x = float(self.x_entry.get())
            y = float(self.y_entry.get())
            self.indexer.set_origin(x, y)
            messagebox.showinfo("Успех", f"Начало координат установлено: X={x}, Y={y}")
            self.status_label.config(text="Начало координат установлено", fg="green")
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректные числовые значения для координат")
    
    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=os.path.basename(file_path), fg="green")
            self.process_btn.config(state="normal")
            self.status_label.config(text="Файл загружен, готов к обработке", fg="green")
    
    def process_file(self):
        if not self.file_path:
            messagebox.showerror("Ошибка", "Сначала загрузите файл")
            return
            
        if self.indexer.origin_x is None or self.indexer.origin_y is None:
            messagebox.showerror("Ошибка", "Сначала установите начало координат")
            return
            
        try:
            self.progress.start()
            self.status_label.config(text="Обработка файла...", fg="blue")
            self.root.update()
            
            # Чтение файла
            df = pd.read_excel(self.file_path)
            
            # Получение названий столбцов
            x_col = self.x_col_entry.get().strip()
            y_col = self.y_col_entry.get().strip()
            
            if x_col not in df.columns or y_col not in df.columns:
                messagebox.showerror("Ошибка", f"Столбцы '{x_col}' и/или '{y_col}' не найдены в файле")
                return
            
            # Обработка данных
            nomenclatural_indices = []
            list_numbers = []
            
            for idx, row in df.iterrows():
                try:
                    x_val = float(row[x_col])
                    y_val = float(row[y_col])
                    
                    nomenclatural_index = self.indexer.calculate_nomenclatural_index(x_val, y_val)
                    list_number = self.indexer.calculate_list_number(x_val, y_val)
                    
                    nomenclatural_indices.append(nomenclatural_index)
                    list_numbers.append(list_number)
                    
                except (ValueError, TypeError) as e:
                    nomenclatural_indices.append("Ошибка")
                    list_numbers.append("Ошибка")
            
            # Добавление результатов в DataFrame
            df['Номенклатурный_индекс'] = nomenclatural_indices
            df['Лист_карты'] = list_numbers
            
            # Сохранение результата
            output_path = filedialog.asksaveasfilename(
                title="Сохранить результат как",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if output_path:
                df.to_excel(output_path, index=False)
                self.status_label.config(text="Файл успешно обработан и сохранен", fg="green")
                messagebox.showinfo("Успех", f"Файл сохранен как:\n{output_path}")
                
                # Показать предпросмотр результата
                self.show_preview(df.head())
            else:
                self.status_label.config(text="Обработка отменена", fg="orange")
                
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при обработке: {str(e)}")
            self.status_label.config(text="Ошибка обработки", fg="red")
        finally:
            self.progress.stop()
    
    def show_preview(self, df_preview):
        """Показывает предпросмотр обработанных данных"""
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Предпросмотр данных")
        preview_window.geometry("700x300")
        
        tk.Label(preview_window, text="Предпросмотр обработанных данных (первые строки)", 
                font=("Arial", 12, "bold")).pack(pady=10)
        
        # Создаем фрейм для таблицы
        frame = tk.Frame(preview_window)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Создаем treeview для отображения данных
        columns = list(df_preview.columns)
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=8)
        
        # Настраиваем заголовки
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)
        
        # Добавляем данные
        for _, row in df_preview.iterrows():
            tree.insert("", "end", values=list(row))
        
        # Добавляем scrollbar
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        tk.Button(preview_window, text="Закрыть", command=preview_window.destroy).pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()