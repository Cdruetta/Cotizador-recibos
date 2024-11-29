import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QFormLayout, QComboBox,
    QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QMessageBox
)
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.units import inch

class CotizacionApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Generador de Cotizaciones')
        self.setGeometry(100, 100, 800, 500)

        self.layout = QVBoxLayout()
        self.form_layout = QFormLayout()

        # Diccionario para almacenar los precios de los productos
        self.productos_precios = {}

        # Menús desplegables
        self.cliente_dropdown = QComboBox()
        self.producto_dropdown = QComboBox()
        self.proveedor_dropdown = QComboBox()

        # Conectar el evento de selección del producto
        # Usamos currentIndexChanged en lugar de activated para capturar la selección correctamente
        self.producto_dropdown.currentIndexChanged.connect(self.actualizar_precio_unitario)

        # Inputs adicionales
        self.cantidad_input = QLineEdit()
        self.precio_input = QLineEdit()
        self.precio_input.setReadOnly(True)  # Hacer que el campo sea de solo lectura

        self.form_layout.addRow('Cliente:', self.cliente_dropdown)
        self.form_layout.addRow('Producto:', self.producto_dropdown)
        self.form_layout.addRow('Proveedor:', self.proveedor_dropdown)
        self.form_layout.addRow('Cantidad:', self.cantidad_input)
        self.form_layout.addRow('Precio Unitario:', self.precio_input)

        # Botón para generar cotización
        self.generar_btn = QPushButton('Generar Cotización')
        self.generar_btn.clicked.connect(self.generar_cotizacion)

        # Tabla para mostrar cotizaciones anteriores
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(['Cliente', 'Producto', 'Proveedor', 'Cantidad', 'Total'])

        self.layout.addLayout(self.form_layout)
        self.layout.addWidget(self.generar_btn)
        self.layout.addWidget(self.table)
        self.setLayout(self.layout)

        # Cargar datos desde el archivo Excel
        self.cargar_datos()

    def cargar_datos(self):
        """
        Carga los datos desde las hojas Clientes, Productos y Proveedores del archivo Excel.
        """
        try:
            # Leer las hojas del archivo Excel
            excel_data = pd.ExcelFile('base_datos.xlsx')

            # Cargar Clientes
            if 'Clientes' in excel_data.sheet_names:
                try:
                    clientes = pd.read_excel(excel_data, sheet_name='Clientes')['Nombre'].dropna().str.strip().unique()
                    self.cliente_dropdown.addItems(clientes.tolist())  # Asegúrate de pasar una lista
                except KeyError:
                    print("La hoja 'Clientes' no contiene una columna llamada 'Nombre'.")

            # Cargar Productos
            if 'Productos' in excel_data.sheet_names:
                try:
                    productos_df = pd.read_excel(excel_data, sheet_name='Productos')

                    # Limpieza de datos
                    productos_df['Nombre'] = productos_df['Nombre'].astype(str).str.strip()
                    productos_df['Precio'] = pd.to_numeric(productos_df['Precio'], errors='coerce')

                    # Filtrar productos con precios válidos
                    productos_df = productos_df.dropna(subset=['Precio'])

                    # Verificar los productos cargados
                    print("Productos cargados:", productos_df[['Nombre', 'Precio']])  # Para depuración

                    # Actualizar el menú desplegable y el diccionario de precios
                    productos = productos_df['Nombre'].unique()
                    self.producto_dropdown.addItems(productos.tolist())  # Asegúrate de pasar una lista
                    self.productos_precios = productos_df.set_index('Nombre')['Precio'].to_dict()
                except KeyError:
                    print("La hoja 'Productos' no contiene las columnas 'Nombre' o 'Precio'.")

            # Cargar Proveedores
            if 'Proveedores' in excel_data.sheet_names:
                try:
                    proveedores = pd.read_excel(excel_data, sheet_name='Proveedores')['Nombre'].dropna().str.strip().unique()
                    self.proveedor_dropdown.addItems(proveedores.tolist())  # Asegúrate de pasar una lista
                except KeyError:
                    print("La hoja 'Proveedores' no contiene una columna llamada 'Nombre'.")

        except FileNotFoundError:
            print("El archivo 'base_datos.xlsx' no fue encontrado. Asegúrate de que esté en el directorio.")

    def actualizar_precio_unitario(self):
        """
        Actualiza el campo de precio unitario según el producto seleccionado.
        """
        producto_seleccionado = self.producto_dropdown.currentText()
        if producto_seleccionado:
            precio = self.productos_precios.get(producto_seleccionado, 0)  # Si no encuentra, devuelve 0
            self.precio_input.setText(str(precio))
        else:
            self.precio_input.clear()

    def generar_cotizacion(self):
        try:
            cliente = self.cliente_dropdown.currentText()
            producto = self.producto_dropdown.currentText()
            proveedor = self.proveedor_dropdown.currentText()

            # Validar entradas numéricas
            if not self.cantidad_input.text().isdigit():
                raise ValueError("La cantidad debe ser un número entero válido.")

            cantidad = int(self.cantidad_input.text())
            precio = float(self.precio_input.text())
            total = cantidad * precio

            # Agregar nueva cotización a la tabla
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            self.table.setItem(row_position, 0, QTableWidgetItem(cliente))
            self.table.setItem(row_position, 1, QTableWidgetItem(producto))
            self.table.setItem(row_position, 2, QTableWidgetItem(proveedor))
            self.table.setItem(row_position, 3, QTableWidgetItem(str(cantidad)))
            self.table.setItem(row_position, 4, QTableWidgetItem(str(total)))

            # Guardar en Excel
            self.guardar_en_excel(cliente, producto, proveedor, cantidad, total)

            # Generar cotización en PDF
            self.generar_pdf(cliente, producto, proveedor, cantidad, precio, total)

        except ValueError as e:
            QMessageBox.warning(self, "Entrada inválida", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Se produjo un error inesperado: {str(e)}")

    def guardar_en_excel(self, cliente, producto, proveedor, cantidad, total):
        """
        Guarda la cotización en un archivo Excel.
        """
        cotizaciones = pd.DataFrame({
            'Cliente': [cliente],
            'Producto': [producto],
            'Proveedor': [proveedor],
            'Cantidad': [cantidad],
            'Total': [total]
        })

        try:
            df = pd.read_excel('base_datos.xlsx', sheet_name='Cotizaciones')
            cotizaciones = pd.concat([df, cotizaciones], ignore_index=True)
        except (FileNotFoundError, ValueError):
            pass  # Si no existe el archivo o la hoja, creamos uno nuevo

        with pd.ExcelWriter('base_datos.xlsx', mode='a', if_sheet_exists='overlay') as writer:
            cotizaciones.to_excel(writer, sheet_name='Cotizaciones', index=False)

    def generar_pdf(self, cliente, producto, proveedor, cantidad, precio, total):
        """
        Genera un archivo PDF con la cotización con un formato profesional.
        """
        try:
            file_name = f'cotizacion_{cliente}_{producto}.pdf'
            document = SimpleDocTemplate(file_name, pagesize=letter)

            # Estilo de texto
            styles = getSampleStyleSheet()
            style_title = styles['Title']
            style_normal = styles['Normal']

            # Crear un lista para los elementos del PDF
            elements = []

            # Encabezado de la cotización
            title = Paragraph(f'<b>Cotización</b><br/>Cliente: {cliente}<br/>Producto: {producto}', style_title)
            elements.append(title)

            # Agregar una línea separadora
            elements.append(Paragraph('<hr />', style_normal))

            # Tabla de productos
            data = [
                ['Producto', 'Proveedor', 'Cantidad', 'Precio Unitario', 'Total'],
                [producto, proveedor, str(cantidad), f"${precio:.2f}", f"${total:.2f}"]
            ]

            # Crear la tabla con los datos
            table = Table(data, colWidths=[2.0*inch, 2.0*inch, 1.0*inch, 1.0*inch, 1.0*inch])
            table.setStyle(TableStyle([ 
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('FONTSIZE', (0, 0), (-1, -1), 12)
            ]))

            elements.append(table)

            # Generar PDF
            document.build(elements)

            QMessageBox.information(self, "PDF Generado", f"Se ha generado el archivo PDF de la cotización: {file_name}")

        except Exception as e:
            QMessageBox.critical(self, "Error al generar PDF", f"Se produjo un error al generar el PDF: {str(e)}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = CotizacionApp()
    window.show()
    sys.exit(app.exec_())
