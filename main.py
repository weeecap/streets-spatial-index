import sys
from PyQt5.QtWidgets import QApplication

from ui import ExcelProcessorApp

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion') 

    window = ExcelProcessorApp()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()