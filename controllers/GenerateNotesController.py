from utils.DataValidator import DataValidator
from utils.CheckExcelProcess import check_excel_running
import os

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class GenerateNotesController:
    """
    Controlador para gestionar la generación de boletines de calificaciones.

    Este controlador se encarga de validar los datos de entrada, verificar que no haya
    instancias de Excel abiertas, y coordinar la creación de carpetas y archivos para
    los boletines de calificaciones.

    Atributos:
        extraction: Objeto que contiene los datos extraídos del archivo Excel.
        save_path (str): Ruta de la carpeta donde se guardarán los boletines.
    """

    def generate_notes(self):
        """
        Genera los boletines de calificaciones.

        Este método realiza las siguientes acciones:
        1. Verifica que no haya instancias de Excel abiertas.
        2. Valida los campos de entrada proporcionados por el usuario.
        3. Valida que se haya seleccionado un archivo y una ruta de guardado.
        4. Extrae los datos necesarios (estudiantes, año escolar y materias) del archivo Excel.
        5. Obtiene los parámetros seleccionados por el usuario.
        6. Crea la ruta de la carpeta donde se guardarán los boletines.
        7. Llama al método `create_folders_boletines` para generar los boletines.

        Excepciones:
            Muestra un mensaje de advertencia si ocurre un error durante el proceso.
        """
        try:
            self.check_excel_running()

            DataValidator.validate_fields(self.view.get_data_inputs())

            self.validate_extraction_and_save_path()

            students, school_year, subjects = self.extraction.get_all_data()
            params = self.view.get_all_data_selected()
            name_folder = params[0].replace('"', "")

            path = os.path.join(self.save_path, name_folder.replace(' ', "_"))
            self.create_folders_boletines(params, students, school_year, subjects, path)

        except Exception as e:
            self.view.show_message(f'Ha ocurrido un error: {e}', 'warning')

    def check_excel_running(self):
        """
        Verifica si hay instancias de Excel abiertas.

        Este método utiliza la función `check_excel_running` para determinar si el programa
        Excel está en ejecución. Si es así, lanza una excepción para evitar conflictos.

        Excepciones:
            Lanza una excepción si se detecta que Excel está abierto.
        """
        if check_excel_running():
            raise Exception("Se ha detectado que el programa Excel se encuentra abierto, por favor ciérrelo para continuar.")

    def validate_extraction_and_save_path(self):
        """
        Valida que se haya seleccionado un archivo y una ruta de guardado.

        Este método verifica que:
        1. Se haya seleccionado un archivo Excel (self.extraction no es None).
        2. Se haya seleccionado una ruta de guardado (self.save_path no está vacío).

        Excepciones:
            Lanza una excepción si no se cumple alguna de las validaciones.
        """
        if self.extraction is None:
            raise Exception('No se ha seleccionado ningún archivo.')

        if not self.save_path:
            raise Exception("No se ha seleccionado una ruta para guardar los boletines.")