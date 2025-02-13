import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QFormLayout, QComboBox, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QMessageBox
)
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import os
from datetime import datetime


def obtener_ruta_archivo(nombre_archivo):
    """Devuelve la ruta correcta del archivo dependiendo del entorno."""
    if getattr(sys, 'frozen', False):  # Ejecutable .exe
        ruta = os.path.join(sys._MEIPASS, nombre_archivo)
    else:
        ruta = os.path.join(os.path.dirname(__file__), nombre_archivo)
    return ruta


class CotizacionApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Generador de Cotizaciones')
        self.setGeometry(100, 100, 800, 500)

        self.layout = QVBoxLayout()
        self.form_layout = QFormLayout()

        # Diccionario para almacenar los precios de los productos
        self.productos_precios = {}

        # Tabla de productos agregados
        self.productos_agregados = []

        # Menús desplegables
        self.cliente_dropdown = QComboBox()
        self.producto_dropdown = QComboBox()
        self.proveedor_dropdown = QComboBox()

        # Inputs adicionales
        self.cantidad_input = QLineEdit()
        self.precio_input = QLineEdit()
        self.precio_input.setReadOnly(True)

        # Añadir al formulario
        self.form_layout.addRow('Cliente:', self.cliente_dropdown)
        self.form_layout.addRow('Producto:', self.producto_dropdown)
        self.form_layout.addRow('Proveedor:', self.proveedor_dropdown)
        self.form_layout.addRow('Cantidad:', self.cantidad_input)
        self.form_layout.addRow('Precio Unitario:', self.precio_input)

        # Botones
        self.agregar_producto_btn = QPushButton('Agregar Producto')
        self.agregar_producto_btn.clicked.connect(self.agregar_producto)

        self.generar_btn = QPushButton('Generar Cotización')
        self.generar_btn.clicked.connect(self.generar_cotizacion)

        # Tabla para mostrar productos agregados
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(['Producto', 'Proveedor', 'Cantidad', 'Precio Unitario', 'Total'])

        # Layouts
        self.layout.addLayout(self.form_layout)
        self.layout.addWidget(self.agregar_producto_btn)
        self.layout.addWidget(self.table)
        self.layout.addWidget(self.generar_btn)
        self.setLayout(self.layout)

        # Cargar datos desde el archivo Excel
        self.cargar_datos()

        # Conectar el cambio de producto para actualizar el precio unitario
        self.producto_dropdown.currentTextChanged.connect(self.actualizar_precio_unitario)

    def cargar_datos(self):
        """Carga datos desde el archivo Excel."""
        try:
            ruta_base_datos = obtener_ruta_archivo('base_datos.xlsx')
            excel_data = pd.ExcelFile(ruta_base_datos)

            # Clientes
            if 'Clientes' in excel_data.sheet_names:
                clientes = pd.read_excel(excel_data, sheet_name='Clientes')['Nombre'].dropna().unique()
                self.cliente_dropdown.addItems(clientes.tolist())

            # Productos
            if 'Productos' in excel_data.sheet_names:
                productos_df = pd.read_excel(excel_data, sheet_name='Productos')
                productos_df = productos_df.dropna(subset=['Nombre', 'Precio'])
                self.productos_precios = productos_df.set_index('Nombre')['Precio'].to_dict()
                self.producto_dropdown.addItems(productos_df['Nombre'].tolist())

            # Proveedores
            if 'Proveedores' in excel_data.sheet_names:
                proveedores = pd.read_excel(excel_data, sheet_name='Proveedores')['Nombre'].dropna().unique()
                self.proveedor_dropdown.addItems(proveedores.tolist())
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo cargar los datos: {e}")

    def agregar_producto(self):
        """Agrega un producto a la lista y la tabla."""
        try:
            producto = self.producto_dropdown.currentText()
            proveedor = self.proveedor_dropdown.currentText()
            cantidad_texto = self.cantidad_input.text()
            precio_texto = self.precio_input.text()

            if not producto or not proveedor or not cantidad_texto or not precio_texto:
                raise ValueError("Todos los campos deben estar completos.")

            cantidad = int(cantidad_texto)
            precio_unitario = float(precio_texto)

            if cantidad <= 0 or precio_unitario <= 0:
                raise ValueError("Cantidad y precio deben ser mayores que cero.")

            total = cantidad * precio_unitario

            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            self.table.setItem(row_position, 0, QTableWidgetItem(producto))
            self.table.setItem(row_position, 1, QTableWidgetItem(proveedor))
            self.table.setItem(row_position, 2, QTableWidgetItem(str(cantidad)))
            self.table.setItem(row_position, 3, QTableWidgetItem(f"{precio_unitario:.2f}"))
            self.table.setItem(row_position, 4, QTableWidgetItem(f"{total:.2f}"))

            self.productos_agregados.append((producto, proveedor, cantidad, precio_unitario, total))
        except ValueError as e:
            QMessageBox.warning(self, "Entrada Inválida", str(e))

    def actualizar_precio_unitario(self):
        producto = self.producto_dropdown.currentText()
        precio = self.productos_precios.get(producto, 0)
        self.precio_input.setText(f"{precio:.2f}")

    def generar_cotizacion(self):
        cliente = self.cliente_dropdown.currentText()
        proveedor = self.proveedor_dropdown.currentText()
        fecha = datetime.now().strftime("%d/%m/%Y")

        if not self.productos_agregados:
            QMessageBox.warning(self, "Sin Productos", "Agrega al menos un producto antes de generar la cotización.")
            return

        try:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "cotizaciones")
            os.makedirs(desktop_path, exist_ok=True)
            file_path = os.path.join(desktop_path, f"cotizacion_{cliente}.pdf")

            document = SimpleDocTemplate(file_path, pagesize=landscape(letter))
            elements = []
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], alignment=1)

            # Definir el estilo para el encabezado
            style_header = ParagraphStyle(
                'HeaderStyle',
                parent=getSampleStyleSheet()['Heading1'],  # Usamos un estilo predefinido como base
                fontSize=14,
                fontName='Helvetica-Bold',
                textColor=colors.black,
                alignment=1,  # Centrado
                spaceAfter=12  # Espacio después del párrafo
            )
            
            # Logo
            logo_path = obtener_ruta_archivo("img/logo.png")
            if os.path.exists(logo_path):
                logo = Image(logo_path, width=100)   # Ajusta el ancho
                logo.drawHeight = (logo.imageHeight * 100) / logo.imageWidth   # Ajusta la altura proporcionalmente
                elements.append(logo)
                
            # Espaciado y encabezado
            elements.append(Spacer(1, 12))
            elements.append(Paragraph("<b>GCinsumos y Servicio Técnico</b>", style_header))


            # Información del cliente y proveedor
            elements.append(Paragraph(f"<b>Cliente:</b> {cliente}", styles['Normal']))
            elements.append(Paragraph(f"<b>Proveedor:</b> {proveedor}", styles['Normal']))
            elements.append(Paragraph(f"<b>Fecha:</b> {fecha}", styles['Normal']))

            elements.append(Spacer(1, 12))  # Espaciado entre la fecha y la tabla

            # Crear la tabla de productos
            data = [['Producto', 'Proveedor', 'Cantidad', 'Precio Unitario', 'Total']]
            for p, prov, c, pu, t in self.productos_agregados:
                data.append([p, prov, c, f"${pu:.2f}", f"${t:.2f}"])

            table = Table(data)
            table.setStyle(TableStyle([ 
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey), 
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ]))
            elements.append(table)

            elements.append(Spacer(1, 12))
            elements.append(Paragraph("Esta cotizacion es valida por 7 dias.", styles['Italic']))

            document.build(elements)
            QMessageBox.information(self, "Cotización Generada", f"Guardado en {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo generar el PDF: {e}")


if __name__ == '__main__':
    from PyQt5.QtWidgets import QApplication
    app = QApplication(sys.argv)
    window = CotizacionApp()
    window.show()
    sys.exit(app.exec_())
