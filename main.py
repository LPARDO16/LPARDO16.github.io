import os
import sys
import pdfplumber
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox

class PDFToExcelApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Exportar PDF a Excel')
        self.setGeometry(100, 100, 400, 200)

        self.label_pdf = QLabel('Seleccionar PDF:', self)
        self.label_pdf.setGeometry(20, 20, 120, 20)

        self.pdf_path = ''
        self.btn_pdf = QPushButton('Elegir PDF', self)
        self.btn_pdf.setGeometry(150, 20, 100, 30)
        self.btn_pdf.clicked.connect(self.choosePDF)

        self.label_name = QLabel('Nombre del archivo Excel:', self)
        self.label_name.setGeometry(20, 60, 160, 20)

        self.excel_name = QLineEdit(self)
        self.excel_name.setGeometry(190, 60, 160, 30)

        self.btn_export = QPushButton('Exportar a Excel', self)
        self.btn_export.setGeometry(20, 100, 150, 30)
        self.btn_export.clicked.connect(self.exportToExcel)

    def choosePDF(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        self.pdf_path, _ = QFileDialog.getOpenFileName(self, 'Seleccionar PDF', '', 'Archivos PDF (*.pdf);;Todos los archivos (*)', options=options)

    def exportToExcel(self):
        if not self.pdf_path:
            QMessageBox.critical(self, 'Error', 'Por favor, seleccione un archivo PDF.')
            return

        if not self.excel_name.text():
            QMessageBox.critical(self, 'Error', 'Por favor, ingrese un nombre para el archivo Excel.')
            return

        try:
            # Abre el archivo PDF
            with pdfplumber.open(self.pdf_path) as pdf:
                # Variables para almacenar los datos adicionales
                ruc_data = None
                comprobante_data = None
                autorizacion_data = None
                fecha_hora_data = None

                # Itera a través de las páginas del PDF
                for page in pdf.pages:
                    # Extrae el texto de la página
                    text = page.extract_text()

                    # Utiliza expresiones regulares para buscar los datos específicos
                    ruc_pattern = r"R\.U\.C\.: (\d{13})"
                    comprobante_pattern = r"COMPROBANTE DE RETENCIÓN\nNo\. (\d+-\d+-\d+)"
                    autorizacion_pattern = r"NÚMERO DE AUTORIZACIÓN\n(\d+)"
                    fecha_hora_pattern = r"FECHA Y HORA DE\n(.+)"

                    # Busca los datos en el texto
                    ruc_match = re.search(ruc_pattern, text)
                    comprobante_match = re.search(comprobante_pattern, text)
                    autorizacion_match = re.search(autorizacion_pattern, text)
                    fecha_hora_match = re.search(fecha_hora_pattern, text)

                    # Almacena los datos encontrados, si están disponibles
                    if ruc_match:
                        ruc_data = ruc_match.group(1)
                    if comprobante_match:
                        comprobante_data = comprobante_match.group(1)
                    if autorizacion_match:
                        autorizacion_data = autorizacion_match.group(1)
                    if fecha_hora_match:
                        fecha_hora_data = fecha_hora_match.group(1)

                # Extraer datos de la tabla
                table_pattern = r"(\d{13})\s+FACTURA\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{4})\s+([\d.]+)\s+([^0-9]+)\s+([\d.]+)\s+([\d.]+)(?:\s+(\d+))?"
                table_matches = re.findall(table_pattern, text)

                # Crear un DataFrame para los datos de la tabla
                data = []
                headers = ["Comprobante", "Fecha de Emisión", "Ejercicio Fiscal", "Base Imponible para la Retención", "Impuesto", "Porcentaje de Retención", "Valor Retenido"]
                for match in table_matches:
                    comprobante = match[0] + match[7] if match[7] else match[0]
                    fecha_emision = match[1]
                    ejercicio_fiscal = match[2]
                    base_imponible = match[3]
                    impuesto = match[4]
                    porcentaje_retencion = match[5]
                    valor_retenido = match[6]
                    data.append([comprobante, fecha_emision, ejercicio_fiscal, base_imponible, impuesto, porcentaje_retencion, valor_retenido])

                df = pd.DataFrame(data, columns=headers)

                # Ruta por defecto donde se guardará el archivo Excel
                excel_filepath = 'C:/Users/Leonel/Desktop/Datos comprobantes retencion'

                # Guardar el archivo Excel
                excel_filename = self.excel_name.text()
                excel_filepath = os.path.join(excel_filepath, excel_filename + '.xlsx')
                workbook = Workbook()
                sheet = workbook.active

                # Agregar los datos adicionales al archivo Excel
                sheet.append(["R.U.C.", ruc_data])
                sheet.append(["Número de Comprobante de Retención", comprobante_data])
                sheet.append(["Número de Autorización", autorizacion_data])
                sheet.append(["Fecha y Hora de Autorización", fecha_hora_data])
                sheet.append([])  # Agregar una fila en blanco como separación
                for row in dataframe_to_rows(df, index=False, header=True):
                    sheet.append(row)

                # Guardar el archivo Excel en la ubicación especificada
                workbook.save(excel_filepath)

            QMessageBox.information(self, 'Éxito', 'Datos exportados a Excel correctamente.')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Error al exportar a Excel: {str(e)}')

def main():
    app = QApplication(sys.argv)
    window = PDFToExcelApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
