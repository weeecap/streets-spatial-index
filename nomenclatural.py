import math
import re
from typing import Tuple

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
    
    def format_street_name(self, street_name: str) -> str:
        parts = street_name.strip().split()
        
        if len(parts) < 2:
            return street_name
        
        street_types = ['УЛ.', 'ПРОСП.', 'ПР.', 'ПЕР.', 'Ш.', 'НАБ.', 'Б-Р', 'БУЛЬВАР', 'ПЛ.', 'ПЛОЩАДЬ']

        if parts[0] in street_types:
            street_type = parts[0]
            name_only = ' '.join(parts[1:])
            return f"{name_only}, {street_type}"
        elif re.match(r'^\d+-й\b', parts[0], re.IGNORECASE ):
            if parts[1] in street_types:
                name_part = ' '.join(parts[2:])
                prefix_part = ' '.join(parts[:2])
                return f"{name_part}, {prefix_part}"
            elif parts[-1] in street_types:
                name = ' '.join(parts[1:-1])
                prefix = f"{parts[0]} {parts[-1]}"
                return f"{name}, {prefix}"      
        elif parts[-1] in street_types:
            return street_name
        else:
            return street_name