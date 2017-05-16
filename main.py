import sys
import locale

from PyQt5.QtWidgets import QApplication, QDialog
from GeneradorMainWindow import GeneradorMainWindow

# Establecemos el locale de nuestro sistema
locale.setlocale(locale.LC_ALL, "")

if __name__ == "__main__":
    APP = QApplication(sys.argv)
    WIN = GeneradorMainWindow()
    WIN.show()
    sys.exit(APP.exec_())
