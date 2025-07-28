import sys
from PyQt5.QtWidgets import QApplication
from Application_GUI import MyWindow


def window():
    app = QApplication(sys.argv)
    
    win = MyWindow()
    win.show()
    
    sys.exit(app.exec())
    
if __name__ == "__main__":
    window()