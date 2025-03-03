from win32com import client
import threading
import pythoncom

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class PDF:
    """
    Clase para convertir archivos de Excel a PDF utilizando la biblioteca `win32com`.

    Attributes:
        input_path (str): Ruta del archivo Excel de entrada.
        output_path (str): Ruta del archivo PDF de salida.
    """

    def __init__(self, input_path='', output_path=''):
        """
        Inicializa la clase PDF.

        Args:
            input_path (str, optional): Ruta del archivo Excel de entrada. Por defecto es ''.
            output_path (str, optional): Ruta del archivo PDF de salida. Por defecto es ''.
        """
        self.input_path = input_path
        self.output_path = output_path
    
    def run(self):
        """
        Convierte el archivo Excel a PDF utilizando la aplicación de Excel.

        Este método abre el archivo Excel, exporta la primera hoja a PDF y cierra la aplicación de Excel.

        Excepciones:
            Muestra un mensaje de advertencia si ocurre un error inesperado durante el proceso.
        """
        pythoncom.CoInitialize()
        try:
            excel = client.Dispatch("Excel.Application")
            sheets = excel.Workbooks.Open(self.input_path)  
            sheet = sheets.Worksheets[0]
            excel.DisplayAlerts = False
            sheet.ExportAsFixedFormat(0, self.output_path)
        except Exception as e:
            raise Exception(f"Ha ocurrido un error al intentar generar los PDFs: {e}")
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()

def run_export_in_background(input_path: str, output_path: str):
    """
    Ejecuta la conversión de Excel a PDF en un hilo separado.

    Args:
        input_path (str): Ruta del archivo Excel de entrada.
        output_path (str): Ruta del archivo PDF de salida.
    """
    pdf = PDF(input_path, output_path)
    thread = threading.Thread(target=pdf.run())
    thread.start()
    thread.join()