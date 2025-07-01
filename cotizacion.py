import sys
import os
from datetime import datetime
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QFormLayout, QComboBox, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QMessageBox
)
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from PIL import Image as PILImage

# Funciones auxiliares
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
        self.setWindowTitle('Generador de Cotizaciones y Recibos')
        self.setGeometry(100, 100, 800, 500)

        # Inicialización de componentes
        self.layout = QVBoxLayout()
        self.form_layout = QFormLayout()
        self.productos_precios = {}
        self.productos_agregados = []

        self.cliente_dropdown = QComboBox()
        self.producto_dropdown = QComboBox()
        self.proveedor_dropdown = QComboBox()
        self.cantidad_input = QLineEdit()
        self.precio_input = QLineEdit()
        self.precio_input.setReadOnly(True)
        
        # Campos nuevos para la dirección, teléfono y localidad
        self.direccion_input = QLineEdit()
        self.direccion_input.setReadOnly(True)
        self.telefono_input = QLineEdit()
        self.telefono_input.setReadOnly(True)
        self.localidad_input = QLineEdit()
        self.localidad_input.setReadOnly(True)

        # Dropdown para elegir tipo de documento (Presupuesto o Recibo)
        self.tipo_documento_dropdown = QComboBox()
        self.tipo_documento_dropdown.addItems(["Presupuesto", "Recibo"])

        # Añadir al formulario
        self.form_layout.addRow('Cliente:', self.cliente_dropdown)
        self.form_layout.addRow('Dirección:', self.direccion_input)
        self.form_layout.addRow('Teléfono:', self.telefono_input)
        self.form_layout.addRow('Localidad:', self.localidad_input)
        self.form_layout.addRow('Producto:', self.producto_dropdown)
        self.form_layout.addRow('Proveedor:', self.proveedor_dropdown)
        self.form_layout.addRow('Cantidad:', self.cantidad_input)
        self.form_layout.addRow('Precio Unitario:', self.precio_input)
        self.form_layout.addRow('Tipo de Documento:', self.tipo_documento_dropdown)

        # Botones
        self.agregar_producto_btn = QPushButton('Agregar Producto')
        self.agregar_producto_btn.clicked.connect(self.agregar_producto)

        self.generar_btn = QPushButton('Generar Documento')
        self.generar_btn.clicked.connect(self.generar_documento)

        self.nuevo_presupuesto_btn = QPushButton('Nuevo Documento')
        self.nuevo_presupuesto_btn.clicked.connect(self.nuevo_presupuesto)

        # Tabla de productos agregados
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(['Producto', 'Proveedor', 'Cantidad', 'Precio Unitario', 'Total'])

        # Layouts
        self.layout.addLayout(self.form_layout)
        self.layout.addWidget(self.agregar_producto_btn)
        self.layout.addWidget(self.generar_btn)
        self.layout.addWidget(self.nuevo_presupuesto_btn)
        self.layout.addWidget(self.table)
        self.setLayout(self.layout)

        # Cargar datos desde el archivo Excel
        self.cargar_datos()

        # Conectar el cambio de producto para actualizar el precio unitario
        self.producto_dropdown.currentTextChanged.connect(self.actualizar_precio_unitario)
        self.cliente_dropdown.currentTextChanged.connect(self.actualizar_datos_cliente)
        
    def obtener_numero_presupuesto(self):
        """Obtiene el siguiente número de presupuesto desde un archivo."""
        try:
            # Ruta donde se almacena el número del presupuesto
            ruta_numero = obtener_ruta_archivo('numero_presupuesto.txt')
            
            # Verificar si el archivo existe y leer el número actual
            if os.path.exists(ruta_numero):
                with open(ruta_numero, 'r') as file:
                    numero = int(file.read().strip())  # Leer y convertir a entero
            else:
                numero = 1  # Si el archivo no existe, empezar desde 1
            
            # Incrementar el número para el siguiente presupuesto
            siguiente_numero = numero + 1
            
            # Guardar el siguiente número para la próxima vez
            with open(ruta_numero, 'w') as file:
                file.write(str(siguiente_numero))
            
            return numero  # Retornar el número actual, antes de incrementarlo
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo obtener el número del presupuesto: {e}")
            return 1  # En caso de error, empezar desde el 1

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

    def actualizar_datos_cliente(self):
        """Actualiza la información del cliente en los campos de texto."""
        cliente = self.cliente_dropdown.currentText()
        if cliente in self.clientes_data:
            datos = self.clientes_data[cliente]
            self.direccion_input.setText(datos.get('Dirección', ''))
            self.telefono_input.setText(str(datos.get('Teléfono', '')))
            self.localidad_input.setText(datos.get('Localidad', ''))

    def actualizar_precio_unitario(self):
        """Actualiza el precio unitario cuando se selecciona un producto."""
        producto = self.producto_dropdown.currentText()
        precio = self.productos_precios.get(producto, 0)
        self.precio_input.setText(f"{precio:.2f}")

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

    def nuevo_presupuesto(self):
        """Reinicia los campos y la tabla para crear un nuevo documento."""
        self.cliente_dropdown.setCurrentIndex(0)
        self.producto_dropdown.setCurrentIndex(0)
        self.proveedor_dropdown.setCurrentIndex(0)
        self.cantidad_input.clear()
        self.precio_input.clear()

        # Limpiar la tabla
        self.table.setRowCount(0)

        # Reiniciar la lista de productos agregados
        self.productos_agregados = []
        QMessageBox.information(self, "Nuevo Documento", "Los datos han sido limpiados, ahora puedes crear un nuevo documento.")

    def generar_documento(self):
        """Genera un presupuesto o recibo según el tipo seleccionado."""
        cliente = self.cliente_dropdown.currentText()
        if not cliente:
            QMessageBox.warning(self, "Cliente", "Debe seleccionar un cliente.")
            return

        if not self.productos_agregados:
            QMessageBox.warning(self, "Productos", "Debe agregar al menos un producto.")
            return

        tipo_documento = self.tipo_documento_dropdown.currentText()
        if tipo_documento == "Presupuesto":
            file_path = self.generar_presupuesto(cliente, self.productos_agregados)
        elif tipo_documento == "Recibo":
            file_path = self.generar_recibo(cliente, self.productos_agregados)

        # Confirmar que se generó correctamente
        QMessageBox.information(self, "Éxito", f"{tipo_documento} generado: {file_path}")

    def generar_presupuesto(self, cliente, productos):
        """Genera el presupuesto en formato PDF con un número incremental."""
        numero_presupuesto = self.obtener_numero_presupuesto()
        fecha = datetime.now().strftime("%d/%m/%Y")
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "presupuestos")
        os.makedirs(desktop_path, exist_ok=True)

        # Ya no escribimos el número aquí porque se hace en obtener_numero_presupuesto()

        file_path = os.path.join(desktop_path, f"presupuesto_{numero_presupuesto}_{cliente}.pdf")

        # Crear el documento PDF
        return self.generar_pdf(cliente, productos, "Presupuesto", numero_presupuesto, file_path)

    def generar_recibo(self, cliente, productos):
        """Genera el recibo en formato PDF."""
        numero_recibo = self.obtener_numero_presupuesto()
        fecha = datetime.now().strftime("%d/%m/%Y")
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "recibos")
        os.makedirs(desktop_path, exist_ok=True)

        # Ya no escribimos el número aquí porque se hace en obtener_numero_presupuesto()

        file_path = os.path.join(desktop_path, f"recibo_{numero_recibo}_{cliente}.pdf")

        # Crear el documento PDF
        return self.generar_pdf(cliente, productos, "Recibo", numero_recibo, file_path)

    def generar_pdf(self, cliente, productos, tipo, numero, file_path):
        """Genera el PDF (presupuesto o recibo)."""
        document = SimpleDocTemplate(file_path, pagesize=landscape(letter))
        elements = []
        styles = getSampleStyleSheet()

        # Estilos personalizados
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('EmpresaStyle', fontSize=14, alignment=1)  # Centrado
        datos_style = ParagraphStyle('DatosStyle', fontSize=10, alignment=1)  # Centrado

        # Verificar la existencia del logo
        logo_path = obtener_ruta_archivo('img/logo.png')
        if not os.path.exists(logo_path):
            print(f"El archivo del logo no se encuentra en: {logo_path}")

        logo = Image(logo_path, width=150, height=100)  # Ajusta el tamaño del logo

        # Nombre de la empresa y datos de contacto
        company_name = "SERVICIOS INFORMÁTICOS"
        datos_contacto = "Dilkendein 1278 - Tel: 358-4268768 - Email: cristian.e.druetta@gmail.com"

        company_name_paragraph = Paragraph(f"<b>{company_name}</b>", title_style)
        datos_contacto_paragraph = Paragraph(f"<i>{datos_contacto}</i>", datos_style) #Italic

        # Tabla con una sola fila y columna para centrar todo
        header_table = Table([[logo, company_name_paragraph, datos_contacto_paragraph]], colWidths=[150, 250, 250])

        header_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ]))

        elements = []
        elements.append(header_table)  # Agregar tabla con logo y texto
        elements.append(Spacer(1, 12))  # Espacio después del encabezado

        # Título del documento (Recibo o Presupuesto)
        elements.append(Paragraph(f"<b>{tipo} N° {numero}</b>", title_style))
        elements.append(Spacer(1, 12))

        # Datos del Cliente
        elements.append(Paragraph(f"<b>Cliente:</b> {cliente}", styles['Normal']))
        elements.append(Paragraph(f"<b>Dirección:</b> {self.direccion_input.text()}", styles['Normal']))
        elements.append(Paragraph(f"<b>Teléfono:</b> {self.telefono_input.text()}", styles['Normal']))
        elements.append(Paragraph(f"<b>Localidad:</b> {self.localidad_input.text()}", styles['Normal']))

        # Fecha
        fecha = datetime.now().strftime("%d/%m/%Y")
        elements.append(Paragraph(f"<b>Fecha:</b> {fecha}", styles['Normal']))
        elements.append(Spacer(1, 12))

        # Tabla de productos
        data = [["Producto", "Cantidad", "Precio Unitario", "Total"]]
        total_general = 0
        for producto, proveedor, cantidad, precio_unitario, total in productos:
            data.append([producto, cantidad, f"${precio_unitario:.2f}", f"${total:.2f}"])
            total_general += total
        
        total_text = f"${total_general:.2f}"
        total_paragraph = Paragraph(f"<b>{total_text}</b>", ParagraphStyle('BoldStyle', fontSize=12, fontName='Helvetica-Bold'))
        data.append(["", "", "Total:", total_paragraph])

        # Ajuste de columnas
        col_widths = [400, 80, 100, 100]
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([ 
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),  
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
        ]))
        
        # Añadir la tabla
        elements.append(table)

        # Footer
        footer_style = ParagraphStyle('FooterStyle', parent=styles['Normal'], alignment=1)
        elements.append(Spacer(1, 12))
        elements.append(Paragraph("Este documento tiene validez por 7 días.", footer_style))
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
