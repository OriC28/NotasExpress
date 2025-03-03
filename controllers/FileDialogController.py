from models.Extraction.ExtractionData import ExtractionData
from models.Extraction.SetExtraction import SetExtraction
from utils.DataValidator import DataValidator

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class FileDialogController:
    """
    Controlador para gestionar los diálogos de selección de archivos y carpetas.

    Este controlador maneja la interacción con el usuario para seleccionar archivos Excel
    y carpetas donde se guardarán los boletines generados. También se encarga de extraer
    y mostrar las hojas de cálculo disponibles en el archivo Excel seleccionado.

    Atributos:
        file_path (str): Ruta del archivo Excel seleccionado.
        save_path (str): Ruta de la carpeta donde se guardarán los boletines.
    """

    def select_excel_file(self):
        """
        Abre un diálogo para seleccionar un archivo Excel y procesa su contenido.

        Este método realiza las siguientes acciones:
        1. Limpia la interfaz gráfica antes de abrir el diálogo de selección de archivo.
        2. Abre un diálogo para que el usuario seleccione un archivo Excel.
        3. Valida que se haya seleccionado un archivo.
        4. Muestra un mensaje de confirmación para generar los boletines.
        5. Si el usuario confirma, extrae los datos del archivo Excel y habilita los widgets
           necesarios en la interfaz gráfica.

        Excepciones:
            Muestra un mensaje de advertencia si ocurre un error inesperado durante el proceso.
        """
        try:
            self.view.clean_all_before_open_file_dialog()
            file_path = self.view.open_file_dialog_to_file()
            if not file_path:
                self.view.show_message('La ruta del archivo xlsx es obligatoria.', 'warning')
            else:
                result = self.view.show_message('¿Está seguro de generar estos boletines?', 'question')
                if(result):
                    setExtraction = SetExtraction(file_path, 0)
                    DataValidator.validate_file_input(setExtraction, self.view)
                    self.file_path = file_path
                    self.view.pathFile.setText(file_path)
                    extraction = ExtractionData(file_path)
                    self.set_sheets_in_gui(extraction)
                    self.view.activate_widgets()
                    del extraction
        except Exception as e:
            self.view.show_message(f'Ha ocurrido un error inesperado al cargar el archivo. {e}', 'warning')

    def select_path_to_save(self):
        """
        Abre un diálogo para seleccionar una carpeta donde guardar los boletines.

        Este método realiza las siguientes acciones:
        1. Abre un diálogo para que el usuario seleccione una carpeta.
        2. Si se selecciona una carpeta, habilita el botón de guardado y actualiza
           la interfaz gráfica con la ruta seleccionada.

        Excepciones:
            Muestra un mensaje de advertencia si ocurre un error inesperado durante el proceso.
        """
        try:
            dir_path = self.view.open_file_dialog_to_save()
            if(dir_path):
                self.view.SaveButton.setEnabled(True)
                self.save_path = dir_path
                self.view.pathFolder.setText(self.save_path)
        except Exception as e:
            self.view.show_message(f'Ha ocurrido un error inesperado al cargar la ruta de destino. {e}', 'warning')

    def set_sheets_in_gui(self, extraction: ExtractionData):
        """
        Filtra y muestra las hojas de cálculo disponibles en la interfaz gráfica.

        Este método filtra las hojas de cálculo que no comienzan con '6to' y las
        muestra en la interfaz gráfica para que el usuario pueda seleccionarlas.

        Args:
            extraction (ExtractionData): Objeto que contiene los datos extraídos del archivo Excel.
        """
        sheets = [sheet for sheet in extraction.sheets if not sheet.startswith('6to')]
        self.view.set_sheets(sheets)