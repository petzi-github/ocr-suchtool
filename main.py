#main.py
from PySide6.QtWidgets import QApplication
from gui import OCRApp

if __name__ == "__main__":
    app = QApplication([])
    fenster = OCRApp()
    fenster.show()
    app.exec()
