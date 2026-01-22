import re
import math
import logging
import pandas as pd

from typing import Tuple
from PyQt5.QtCore import QThread, pyqtSignal


class NomenclaturalStreetIndexer:
    def __init__(self, square_size: int = 500):
        self.square_size = square_size
        self.origin_x = None
        self.origin_y = None
        
        self.letters = [chr(i) for i in range(1040, 1072)]
        self.letters = [letter for letter in self.letters if letter != 'Ё' and letter != 'Й' and letter != 'Ы' and letter != 'Ь' and letter != 'Ъ']
        
    def set_origin(self, x: float, y: float) -> None:
        self.origin_x = x
        self.origin_y = y

    def _calculate_indices(self, x: float, y: float) -> Tuple[float, float, int, int]:
        if self.origin_x is None or self.origin_y is None:
            raise ValueError("Начало координат не установлено!")
        
        delta_x = self.origin_x - x 
        delta_y = y - self.origin_y
        
        col_index = math.floor(delta_x / self.square_size)
        row_index = math.floor(delta_y / self.square_size)
        
        return delta_x, delta_y, col_index, row_index
    
    def calculate_nomenclatural_index(self, x: float, y: float) -> str:

        delta_x, delta_y, col_index, row_index = self._calculate_indices(x, y)
        
        letter = self.letters[col_index % len(self.letters)]
        number = row_index + 1
        
        return f"{letter}-{number}"
    
    def calculate_list_number(self, x: float, y: float, border: int = 4500) -> str:
        delta_x, delta_y, col_index, row_index = self._calculate_indices(x, y)
        number = row_index + 1
        
        if number <= 9:
            list_number = "Лист 1" if delta_x <= border else "Лист 3"
        else:
            list_number = "Лист 2" if delta_x <= border else "Лист 4"

        return list_number
    
    # def capitalize(self, street_name: str) -> str: 
    #     street_capitalizating = street_name.capitalize()

    #     if '.' in street_capitalizating:
    #         str


    #     return None 
    
    def format_street_name(self, street_name: str) -> str:
        
        parts = street_name.strip().split()
        
        if len(parts) < 2:
            return street_name
        
        street_types = ['УЛ.', 'ПРОСП.', 'ПР.', 'ПЕР.', 'Ш.', 'НАБ.', 'Б-Р', 'БУЛЬВАР', 'ПЛ.', 'ПЛОЩАДЬ']

        if parts[0] in street_types:
            street_type = parts[0].lower()
            name_only = ' '.join(parts[1:]).capitalize()
            return f"{name_only}, {street_type}"
        elif re.match(r'^\d+-й\b', parts[0], re.IGNORECASE ):
            if parts[1] in street_types:
                name_part = ' '.join(parts[2:]).capitalize()
                prefix_part = ' '.join(parts[:2]).lower()
                return f"{name_part}, {prefix_part}"
            elif parts[-1] in street_types:
                name = ' '.join(parts[1:-1]).capitalize()
                prefix = f"{parts[0].lower()} {parts[-1].lower()}"
                return f"{name}, {prefix}"      
        elif parts[-1] in street_types:
            return street_name.capitalize()
        else:
            return street_name.capitalize()
     
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
            df = pd.read_excel(self.file_path)
            
            if self.x_col not in df.columns or self.y_col not in df.columns:
                self.error_occurred.emit(f"Столбцы '{self.x_col}' и/или '{self.y_col}' не найдены в файле")
                logging.critical('Column X or Y not found')
                return
            
            if self.street_col and self.street_col not in df.columns:
                self.error_occurred.emit(f"Столбец '{self.street_col}' не найден в файле")
                logging.critical('Column Street not found')
                return
            
            street_indices = {}
            street_occurrences = {}

            nomenclatural_indices = []
            final_indices = []
            list_numbers = []
            formatted_streets = []
            total_rows = len(df)
            
            for idx, row in df.iterrows():
                try:
                    x_val = float(row[self.x_col])
                    y_val = float(row[self.y_col])
                    
                    nomenclatural_index = self.indexer.calculate_nomenclatural_index(x_val, y_val)
                    list_number = self.indexer.calculate_list_number(x_val, y_val)
                    logging.info('self calculation passed')

                    if self.street_col:
                        street_name = str(row[self.street_col])

                        if street_name in street_indices:
                            street_indices[street_name].add(nomenclatural_index)
                            street_occurrences[street_name] += 1
                        else:
                            street_indices[street_name] = {nomenclatural_index}
                            street_occurrences[street_name] = 1
                    
                        formatted_street = self.indexer.format_street_name(street_name)
                    else:
                        formatted_street = ""
                        street_name = ""
                    logging.info('formating passed')

                    
                    nomenclatural_indices.append(nomenclatural_index)
                    list_numbers.append(list_number)
                    formatted_streets.append(formatted_street)                           

                    
                except (ValueError, TypeError):
                    nomenclatural_indices.append("Ошибка")
                    list_numbers.append("Ошибка")
                    formatted_streets.append("Ошибка")
                    logging.critical('function raised error')

                progress = int((idx + 1) / total_rows * 100)
                self.progress_updated.emit(progress)
            
            logging.info(f"Starting second pass. Total rows: {len(df)}")
            for idx, row in df.iterrows():
                try:
                    if self.street_col:
                        street_name = str(row[self.street_col])
                        
                        if street_name in street_occurrences:
                            logging.info(f"street_indices[{street_name}] = {street_indices[street_name]}, type = {type(street_indices[street_name])}")
                            sorted_indices = sorted(list(street_indices[street_name]))

                            first_index_chars = list(sorted_indices[0])

                            number_parts = []
                            for idx in sorted_indices:
                                for i, char in enumerate(idx):
                                    if char.isdigit():
                                        number_part = idx[i:]  
                                        number_parts.append(number_part)
                                        break
                                else:
                                    number_parts.append(idx)

                            prefix_end = 0
                            for i, char in enumerate(first_index_chars):
                                if char.isdigit():
                                    prefix_end = i
                                    break

                            prefix = ''.join(first_index_chars[:prefix_end])

                            if all(idx.startswith(prefix) for idx in sorted_indices):
                                numbers = [idx[len(prefix):] for idx in sorted_indices]
                                
                                if len(numbers) == 1:
                                    final_index = f"{prefix}{numbers[0]}"
                                else:
                                    final_index = f"{prefix}{numbers[0]}, {', '.join(numbers[1:])}"

                            elif len(number_parts) >= 2:
                                number_to_prefixes = {}
                                for idx, num_part in zip(sorted_indices, number_parts):
                                    if num_part not in number_to_prefixes:
                                        number_to_prefixes[num_part] = []
                                    
                                    prefix_chars = []
                                    for char in idx:
                                        if char.isdigit():
                                            break
                                        if char != '-':
                                            prefix_chars.append(char)
                                    prefix_str = ''.join(prefix_chars)
                                    
                                    number_to_prefixes[num_part].append(prefix_str)
                                
                                result_parts = []
                                for num_part, prefixes in number_to_prefixes.items():
                                    if len(prefixes) == 1:
                                        result_parts.append(f"{prefixes[0]}-{num_part}")
                                    else:
                                        prefixes_str = ', '.join(prefixes)
                                        result_parts.append(f"{prefixes_str}-{num_part}")
                                
                                final_index = "; ".join(result_parts)
                            else:
                                final_index = nomenclatural_indices[idx]
                        else:
                            final_index = nomenclatural_indices[idx]
                    else:
                        final_index = nomenclatural_indices[idx]
                    
                    final_indices.append(final_index)
                    logging.info('indices calculation passed')
                
                except Exception as e:
                    logging.error(f"Error in second pass at row {idx}: {e}")
                    final_indices.append(nomenclatural_indices[idx])

            result_df = pd.DataFrame()

            status_list = []
            for idx in range(len(df)):
                logging.info(f'{idx} is unique')  
                street_name = str(df.iloc[idx][self.street_col]).strip()
                if street_name and street_name in street_occurrences and street_occurrences[street_name] > 1:
                    status_list.append("Повторяется")
                else:
                    status_list.append('Уникальное')   
            
            result_df['Номенклатурный индекс'] = final_indices
            result_df['Лист карты'] = list_numbers
            result_df['Форматированная улица'] = formatted_streets
            result_df['Статус уникальности'] = status_list

            self.df_result = result_df
            result_df.drop_duplicates(subset=['Форматированная улица', 'Номенклатурный индекс'], inplace=True, keep='first')
            result_df.sort_values(by = 'Форматированная улица', inplace=True)
            self.finished_processing.emit(result_df)

        except Exception as e:
            logging.critical('Exception raised')
            self.error_occurred.emit(str(e))