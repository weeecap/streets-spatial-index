import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout

from ui import ExcelProcessorApp, CheckAndMatch

class MainApp(QMainWindow):
    def __init__(self, hmap=None):
        super().__init__()
        self.hmap = hmap  
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("ГИС Белгеодезия")
        self.setGeometry(100, 100, 900, 700)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(0, 0, 0, 0)
        
        self.tab_widget = QTabWidget()
        
        self.excel_processor = ExcelProcessorApp()
        self.check_and_match = CheckAndMatch(hmap=self.hmap)

        self.tab_widget.addTab(self.excel_processor, "📊 Обработка номенклатурных индексов")
        self.tab_widget.addTab(self.check_and_match, "🗺️ Считка объектов")
        
        layout.addWidget(self.tab_widget)
        
    def closeEvent(self, event):
        """При закрытии окна чистим ресурсы обоих инструментов"""
        if hasattr(self.check_and_match, 'cleanup'):
            self.check_and_match.cleanup()
        
        if hasattr(self.excel_processor, 'cleanup'):
            self.excel_processor.cleanup()
            
        event.accept()


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Получаем handle карты из системы (если нужно)
    # hmap = получить_handle_карты()
    hmap = None  # Заглушка, нужно будет заменить на реальный handle
    
    window = MainApp(hmap=hmap)
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()