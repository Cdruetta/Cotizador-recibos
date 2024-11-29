import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QComboBox, QCheckBox, QLineEdit, QFormLayout
from PyQt5.QtCore import Qt
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


# Cargar el archivo de Excel
clientes_df = pd.read_excel('base_datos.xlsx', sheet_name='Clientes')
productos_df = pd.read_excel('base_datos.xlsx', sheet_name='Productos')
proveedores_df = pd.read_excel('base_datos.xlsx', sheet_name='Proveedores')

class PresupuestoApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Generador de Presupuestos y Recibos')
        self.setGeometry(100, 100, 500, 400)

        self.layout = QVBoxLayout()

        # Título
        self.titulo = QLabel("Generar Presupuesto o Recibo")
        self.layout.addWidget(self.titulo)

        # ComboBox para seleccionar el cliente
        self.cliente_combo = QComboBox()
        self.cliente_combo.addItems(clientes_df['Nombre'].tolist())
        self.layout.addWidget(QLabel('Seleccione un cliente:'))
        self.layout.addWidget(self.cliente_combo)

        # ComboBox para seleccionar el proveedor
        self.proveedor_combo = QComboBox()
        self.proveedor_combo.addItems(proveedores_df['Nombre'].tolist())
        self.layout.addWidget(QLabel('Seleccione un proveedor:'))
        self.layout.addWidget(self.proveedor_combo)

        # Checkboxes para los productos
        self.productos_layout = QVBoxLayout()
        self.product_checkboxes = []
        for index, row in productos_df.iterrows():
            checkbox = QCheckBox(f"{row['Descripción']} - ${row['Precio']}")
            self.product_checkboxes.append(checkbox)
            self.productos_layout.addWidget(checkbox)
        self.layout.addWidget(QLabel('Seleccione productos:'))
        self.layout.addLayout(self.productos_layout)

        # Botón para generar presupuesto
        self.generar_button = QPushButton('Generar Presupuesto')
        self.generar_button.clicked.connect(self.generar_presupuesto)
        self.layout.addWidget(self.generar_button)

        self.setLayout(self.layout)

    def generar_presupuesto(self):
        cliente_seleccionado = self.cliente_combo.currentText()
        proveedor_seleccionado = self.proveedor_combo.currentText()
        productos_seleccionados = [checkbox.text().split(' - $')[0] for checkbox in self.product_checkboxes if checkbox.isChecked()]
        precios = [float(checkbox.text().split(' - $')[1]) for checkbox in self.product_checkboxes if checkbox.isChecked()]
        total = sum(precios)

        # Mostrar el resumen
        resumen = f"Cliente: {cliente_seleccionado}\n"
        resumen += f"Proveedor: {proveedor_seleccionado}\n"
        resumen += "Productos seleccionados:\n"
        for producto in productos_seleccionados:
            resumen += f"  - {producto}\n"
        resumen += f"Total: ${total:.2f}"

        # Crear un archivo PDF
        self.generar_pdf(cliente_seleccionado, proveedor_seleccionado, productos_seleccionados, total)

        print(resumen)

    def generar_pdf(self, cliente, proveedor, productos, total):
        # Guardar el presupuesto o recibo como PDF
        filename = f"{cliente}_presupuesto.pdf"
        c = canvas.Canvas(filename, pagesize=letter)
        c.drawString(100, 750, f"Presupuesto para {cliente}")
        c.drawString(100, 730, f"Proveedor: {proveedor}")
        y_position = 710
        for producto in productos:
            c.drawString(100, y_position, f"- {producto}")
            y_position -= 20
        c.drawString(100, y_position, f"Total: ${total:.2f}")
        c.save()

        print(f"Presupuesto generado y guardado como {filename}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PresupuestoApp()
    window.show()
    sys.exit(app.exec_())
