import sys
import os
from PyQt5.QtWidgets import QApplication
from cotizacion import CotizacionApp

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = CotizacionApp()
    window.show()
    sys.exit(app.exec_())
