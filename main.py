import math

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
        # Исключаем Ё
        self.letters = [letter for letter in self.letters if letter != 'Ё' and letter != 'Й']
        
    def set_origin(self, x: float, y: float):
        
        self.origin_x = x
        self.origin_y = y
        print(f"Начало координат установлено: X={x}, Y={y}")
    
    def calculate_nomenclatural_index(self, x: float, y: float) -> str:
        """
        
        Args:
            x: координата X
            y: координата Y
            
        Returns:
            Номенклатурный индекс (строка)
        """
        if self.origin_x is None or self.origin_y is None:
            raise ValueError("Начало координат не установлено!")
        
        delta_x = self.origin_x - x 
        delta_y = y - self.origin_y
        
        col_index = math.floor(delta_x / self.square_size)
        row_index = math.floor(delta_y / self.square_size)
        
        letter = self.letters[col_index % len(self.letters)]
        number = row_index + 1  # Нумерация обычно начинается с 1
        
        return f"{letter}-{number}"

if __name__ == "__main__":
    indexer = NomenclaturalStreetIndexer(500)
    
    indexer.set_origin(5946000, 5417000)
    
    x = 5940191.697
    y = 5422057.239
    nomenclatural_index = indexer.calculate_nomenclatural_index(x, y)
    
    print(f"Координаты: X={x}, Y={y}")
    print(f"Номенклатурный индекс: {nomenclatural_index}")