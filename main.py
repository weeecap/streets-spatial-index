import math
import pandas as pd
import tkinter as tk

from typing import List, Dict, Tuple 

class NomenclaturalStreetIndexer:
    def __init__(self, square_size: int = 500):

        self.square_size = square_size
        self.origin_x = None
        self.origin_y = None
        
        self.letters = [chr(i) for i in range(1040, 1072)]  # А-Я
        self.letters = [letter for letter in self.letters if letter != 'Ё' and letter != 'Й']
        
    def set_origin(self, x: float, y: float):
        
        self.origin_x = x
        self.origin_y = y
        print(f"Начало координат установлено: X={x}, Y={y}")

    def _calculate_(self, x: float, y: float) -> tuple:

        if self.origin_x is None or self.origin_y is None:
            raise ValueError("Начало координат не установлено!")
        
        delta_x = self.origin_x - x 
        delta_y = y - self.origin_y
        
        col_index = math.floor(delta_x / self.square_size)
        row_index = math.floor(delta_y / self.square_size)
        
        return delta_x, delta_y, col_index, row_index
    
    def calculate_nomenclatural_index(self, x: float, y: float) -> str:

        delta_x, delta_y, col_index, row_index = self._calculate_(x,y)

        col_index = math.floor(delta_x / self.square_size)
        row_index = math.floor(delta_y / self.square_size)

        letter = self.letters[col_index % len(self.letters)]
        number = row_index + 1  
        
        return f"{letter}-{number}"
    
    def format_street_name(self, street_name:str) -> str:

        parts = street_name.strip().split()
        street_type = parts[0]
        name_only = ' '.join(parts[1:])  
        adjectives = ['ая', 'ый', 'ой', 'ий', 'ого', 'его', 'ему', 'ому', 'им', 'ым', 'ем', 'ом', 'яя', 'ей', 'ую', 'юю', 'ое', 'ее', 'омо', 'ег']

        if name_only[-2:] in adjectives:
            return f"{name_only} {street_type}"
        else:
            return f"{name_only}, {street_type}"

    
    def calculate_list_number(self, x: float, y: float, border: int = 4500) -> str:

        delta_x, delta_y, col_index, row_index = self._calculate_(x,y)
        number = row_index + 1
        
        if number <=9:
            list_number = "Лист 1" if delta_x <= border else "Лист 3"
        else:
            list_number = "Лист 2" if delta_x <= border else "Лист 4"

        return list_number

    def export_to_excel(self) -> str:
        return 'Файл был сохранен'          


if __name__ == "__main__":
    indexer = NomenclaturalStreetIndexer(500)

    print('Введите начало координат')
    x_origin, y_origin = int(input()), int(input())

    indexer.set_origin(x_origin, y_origin)
    
    test= 'ул Якуба Коласа'
    x = 5943402.306
    y = 5421679.414
    format_street = indexer.format_street_name(test) 
    nomenclatural_index = indexer.calculate_nomenclatural_index(x, y)
    list_number = indexer.calculate_list_number(x, y)
    
    print(f"Координаты: X={x}, Y={y}")
    print(f"Номенклатурный индекс: {nomenclatural_index}")
    print(f"Номер листа: {list_number}")
    print(f'Форматированное название: {format_street}')