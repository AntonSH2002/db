import sys
from PyQt5 import QtWidgets
from window import Window


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = Window()
    w.show()
    app.exec_()
