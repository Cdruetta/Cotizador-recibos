import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QFormLayout, QComboBox, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QMessageBox
)
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import os
from datetime import datetime

def obtener_ruta_archivo(nombre_archivo):
    """
    Devuelve la ruta correcta del archivo dependiendo de si el script está siendo ejecutado como un archivo .exe
    o desde el código fuente.
    """
    if getattr(sys, 'frozen', False):  # Si está corriendo como ejecutable
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

        # Menús desplegables
        self.cliente_dropdown = QComboBox()
        self.producto_dropdown = QComboBox()
        self.proveedor_dropdown = QComboBox()

        # Conectar el evento de selección del producto
        self.producto_dropdown.activated.connect(self.actualizar_precio_unitario)

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
        try:
            # Ruta correcta para el archivo base_datos.xlsx
            ruta_base_datos = obtener_ruta_archivo('base_datos.xlsx')
            excel_data = pd.ExcelFile(ruta_base_datos)

            # Cargar Clientes
            if 'Clientes' in excel_data.sheet_names:
                clientes = pd.read_excel(excel_data, sheet_name='Clientes')['Nombre'].dropna().str.strip().unique()
                if clientes.size > 0:
                    self.cliente_dropdown.addItems(clientes.tolist())
                else:
                    print("No se encontraron clientes en la hoja 'Clientes'.")

            # Cargar Productos
            if 'Productos' in excel_data.sheet_names:
                productos_df = pd.read_excel(excel_data, sheet_name='Productos')
                if 'Nombre' in productos_df.columns and 'Precio' in productos_df.columns:
                    productos_df['Nombre'] = productos_df['Nombre'].astype(str).str.strip()
                    productos_df['Precio'] = pd.to_numeric(productos_df['Precio'], errors='coerce')

                    # Filtrar productos con precios válidos
                    productos_df = productos_df.dropna(subset=['Precio'])
                    if not productos_df.empty:
                        productos = productos_df['Nombre'].unique()
                        self.producto_dropdown.addItems(productos.tolist())
                        self.productos_precios = productos_df.set_index('Nombre')['Precio'].to_dict()
                    else:
                        print("No hay productos válidos con precios en la hoja 'Productos'.")
                else:
                    print("La hoja 'Productos' no contiene las columnas 'Nombre' o 'Precio'.")
            else:
                print("No se encontró una hoja llamada 'Productos'.")

            # Cargar Proveedores
            if 'Proveedores' in excel_data.sheet_names:
                proveedores = pd.read_excel(excel_data, sheet_name='Proveedores')['Nombre'].dropna().str.strip().unique()
                if proveedores.size > 0:
                    self.proveedor_dropdown.addItems(proveedores.tolist())
                else:
                    print("No se encontraron proveedores en la hoja 'Proveedores'.")
        except FileNotFoundError:
            print("El archivo 'base_datos.xlsx' no fue encontrado. Asegúrate de que esté en el directorio.")
        except Exception as e:
            print(f"Se produjo un error al cargar los datos: {e}")

    def actualizar_precio_unitario(self):
        """
        Actualiza el campo de precio unitario según el producto seleccionado.
        """
        producto_seleccionado = self.producto_dropdown.currentText()
        if producto_seleccionado:
            precio = self.productos_precios.get(producto_seleccionado, 0)
            self.precio_input.setText(f"{precio:.2f}")
        else:
            self.precio_input.clear()

    def generar_cotizacion(self):
        """
        Genera una cotización y la muestra en la tabla, además de guardarla en el archivo Excel.
        """
        try:
            cliente = self.cliente_dropdown.currentText()
            producto = self.producto_dropdown.currentText()
            proveedor = self.proveedor_dropdown.currentText()

            if not self.cantidad_input.text().isdigit():
                raise ValueError("La cantidad debe ser un número entero válido.")

            cantidad = int(self.cantidad_input.text())
            precio = float(self.precio_input.text())
            total = cantidad * precio

            # Insertar en la tabla de la interfaz gráfica
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            self.table.setItem(row_position, 0, QTableWidgetItem(cliente))
            self.table.setItem(row_position, 1, QTableWidgetItem(producto))
            self.table.setItem(row_position, 2, QTableWidgetItem(proveedor))
            self.table.setItem(row_position, 3, QTableWidgetItem(str(cantidad)))
            self.table.setItem(row_position, 4, QTableWidgetItem(f"{total:.2f}"))

            # Generar el PDF
            self.generar_pdf(cliente, producto, proveedor, cantidad, precio, total)

        except ValueError as e:
            QMessageBox.warning(self, "Entrada inválida", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Se produjo un error inesperado: {e}")

    def generar_pdf(self, cliente, producto, proveedor, cantidad, precio, total):
        """
        Genera un archivo PDF con formato profesional para la cotización, incluyendo un logo.
        """
        try:
            # Obtener la ruta del escritorio y crear la carpeta "cotizaciones"
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            cotizaciones_path = os.path.join(desktop_path, "cotizaciones")
            os.makedirs(cotizaciones_path, exist_ok=True)

            # Crear la ruta completa del archivo PDF
            file_name = f'cotizacion_{cliente}_{producto}.pdf'
            file_path = os.path.join(cotizaciones_path, file_name)
            print(f"Generando el PDF: {file_path}")

            # Crear el documento PDF
            document = SimpleDocTemplate(file_path, pagesize=letter)

            # Estilos
            styles = getSampleStyleSheet()
            style_header = ParagraphStyle(
                name='Header',
                fontSize=24,
                alignment=1,  # Centrado
                spaceAfter=10,
                fontName='Helvetica-Bold',
                textColor=colors.darkblue
            )
            style_normal = ParagraphStyle(name='Normal', fontSize=10, spaceAfter=5)

            elements = []

            # Agregar el logo
            try:
                logo_path = 'img/logo.png'  # Ruta relativa al directorio del script
                if os.path.exists(logo_path):
                    logo = Image(logo_path)
                    logo_width = 100  # Ancho deseado del logo
                    logo_height = (logo_width / logo.imageWidth) * logo.imageHeight
                    logo.drawHeight = logo_height
                    logo.drawWidth = logo_width
                    logo.hAlign = 'LEFT'
                    elements.append(logo)
                else:
                    print(f"El archivo de logo no se encuentra en la ruta: {logo_path}")
            except Exception as e:
                print(f"Error al agregar el logo: {e}")

            # Espaciado y encabezado
            elements.append(Spacer(1, 12))
            elements.append(Paragraph("<b>GCinsumos y Servicio Técnico</b>", style_header))

            # Fecha de la cotización
            fecha = datetime.now().strftime("%d/%m/%Y")
            elements.append(Spacer(1, 6))
            elements.append(Paragraph(f"<b>Fecha:</b> {fecha}", style_normal))

            # Información del cliente
            cliente_info = f"""
            <b>CLIENTE:</b> {cliente}<br/>
            <b>PROVEEDOR:</b> {proveedor}<br/>
            """
            elements.append(Spacer(1, 12))
            elements.append(Paragraph(cliente_info, style_normal))

            # Nota
            elements.append(
                Paragraph(
                    "<b>NOTA:</b> Esta cotización tiene una vigencia de 5 días a partir de la fecha de emisión.",
                    style_normal
                )
            )
            elements.append(Spacer(1, 12))

            # Tabla de productos
            data = [
                ['N°', 'U.D.M.', 'CANTIDAD', 'DESCRIPCIÓN DEL ARTÍCULO', 'PRECIO UNITARIO', 'DESCUENTO', 'IMPORTE'],
                ['1', 'PIEZAS', str(cantidad), producto, f"${precio:.2f}", '0%', f"${total:.2f}"],
            ]

            table = Table(data, colWidths=[40, 60, 60, 200, 80, 80, 80])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            elements.append(table)

            # Guardar el PDF
            document.build(elements)
            print(f"PDF generado con éxito en: {file_path}")

        except Exception as e:
            print(f"Error al generar el PDF: {e}")
