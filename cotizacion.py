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
from PIL import Image as PILImage


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
                clientes_df = pd.read_excel(excel_data, sheet_name='Clientes')
                clientes_df = clientes_df.dropna(subset=['Nombre'])  # Elimina filas sin nombre
                self.clientes_data = clientes_df.set_index('Nombre').to_dict(orient='index')
                self.cliente_dropdown.addItems(clientes_df['Nombre'].tolist())

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

    def obtener_numero_presupuesto(self):
        """Obtiene el próximo número de presupuesto disponible."""
        try:
            ruta_numero = obtener_ruta_archivo('numero_presupuesto.txt')
            if os.path.exists(ruta_numero):
                with open(ruta_numero, 'r') as file:
                    numero = int(file.read().strip())
            else:
                numero = 0  # Si el archivo no existe, comenzamos desde el 0
            return numero + 1
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo leer el número de presupuesto: {e}")
            return 1

    def generar_cotizacion(self):
        """Genera la cotización en formato PDF y la guarda en el escritorio."""
        cliente = self.cliente_dropdown.currentText()
        if not cliente:
            QMessageBox.warning(self, "Cliente", "Debe seleccionar un cliente.")
            return

        if not self.productos_agregados:
            QMessageBox.warning(self, "Productos", "Debe agregar al menos un producto.")
            return

        # Generar el presupuesto (PDF)
        file_path = self.generar_presupuesto(cliente, self.productos_agregados)
        
        # Confirmar que se generó correctamente
        QMessageBox.information(self, "Éxito", f"Cotización generada: {file_path}")

    def generar_presupuesto(self, cliente, productos):
        """Genera el presupuesto en formato PDF con un número incremental."""
        numero_presupuesto = self.obtener_numero_presupuesto()
        fecha = datetime.now().strftime("%d/%m/%Y")
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "presupuestos")
        os.makedirs(desktop_path, exist_ok=True)
        
        # Guardar el número del presupuesto para el siguiente
        try:
            ruta_numero = obtener_ruta_archivo('numero_presupuesto.txt')
            with open(ruta_numero, 'w') as file:
                file.write(str(numero_presupuesto))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo actualizar el número de presupuesto: {e}")

        file_path = os.path.join(desktop_path, f"presupuesto_{numero_presupuesto}_{cliente}.pdf")
        
        document = SimpleDocTemplate(file_path, pagesize=landscape(letter))
        elements = []
        styles = getSampleStyleSheet()

        # Estilos personalizados
        title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=18, alignment=1)
        info_style = ParagraphStyle('InfoStyle', parent=styles['Normal'], fontSize=10)

        # Logo
        logo_path = "img/logo.png"
        if os.path.exists(logo_path):
            pil_img = PILImage.open(logo_path)
            width, height = pil_img.size

            logo = Image(logo_path)
            logo.drawHeight = 50
            logo.drawWidth = width * (logo.drawHeight / height)
            elements.append(logo)

        # Encabezado
        elements.append(Paragraph("<b>SERVICIOS INFORMÁTICOS GC</b>", title_style))
        elements.append(Paragraph("Dilkendein 1278 - Tel: 358-4268768 - Email: cristian.e.druetta@gmail.com", info_style))
        elements.append(Spacer(1, 12))

        # Datos del Cliente
        elements.append(Paragraph(f"<b>Cliente:</b> {cliente}", styles['Normal']))

        # Verificar si el cliente tiene datos adicionales
        if cliente in self.clientes_data:
            cliente_data = self.clientes_data[cliente]
            direccion = cliente_data.get('Dirección', 'No disponible')
            telefono = cliente_data.get('Teléfono', 'No disponible')
            equipo = cliente_data.get('Equipo', 'No disponible')

            # Mostrar la dirección, teléfono y equipo
            elements.append(Paragraph(f"<b>Dirección:</b> {direccion}", styles['Normal']))
            elements.append(Paragraph(f"<b>Teléfono:</b> {telefono}", styles['Normal']))
            elements.append(Paragraph(f"<b>Equipo:</b> {equipo}", styles['Normal']))
        else:
            elements.append(Paragraph("<b>Dirección:</b> Información no disponible", styles['Normal']))
            elements.append(Paragraph("<b>Teléfono:</b> Información no disponible", styles['Normal']))
            elements.append(Paragraph("<b>Equipo:</b> Información no disponible", styles['Normal']))

        elements.append(Paragraph(f"<b>Fecha:</b> {fecha}", styles['Normal']))
        elements.append(Spacer(1, 12))

        # Tabla de productos
        data = [["Producto", "Cantidad", "Precio Unitario", "Total"]]
        total_general = 0
        for producto, proveedor, cantidad, precio_unitario, total in productos:
            data.append([producto, cantidad, f"${precio_unitario:.2f}", f"${total:.2f}"])
            total_general += total

        data.append(["", "", "Total:", f"${total_general:.2f}"])

        # Ajustar los anchos de las columnas para evitar el estiramiento
        col_widths = [400, 80, 100, 100]  # Puedes ajustar estos valores a tu gusto
        table = Table(data, colWidths=col_widths)

        # Agregar el estilo a la tabla
        table.setStyle(TableStyle([ 
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),  
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),           
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),   
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),   
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),     
            ('FONTSIZE', (0, 0), (-1, -1), 10),              
        ]))

        elements.append(table)
        document.build(elements)
        return file_path

        # Agregar la tabla a los elementos
        elements.append(table)

        # Estilo centrado para el footer
        footer_style = ParagraphStyle('FooterStyle', parent=styles['Normal'], alignment=1)  # alignment=1 es para centrar

        # Footer con validez y derechos de autor
        elements.append(Spacer(1, 12))
        elements.append(Paragraph("Este presupuesto tiene validez por 7 días.", footer_style))
        elements.append(Paragraph("© GCsoft-2025. Todos los derechos reservados.", footer_style))

        # Crear el PDF
        document.build(elements)

        return file_path


if __name__ == '__main__':
    from PyQt5.QtWidgets import QApplication
    app = QApplication(sys.argv)
    window = CotizacionApp()
    window.show()
    sys.exit(app.exec_())
