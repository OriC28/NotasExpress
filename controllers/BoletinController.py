from views.MainView import MainView
from views.LoadingView import LoadingView
from models.Writter.Write import Write
from models.Writter.CreateFiles import CreateFiles
from controllers.GenerateNotesController import GenerateNotesController
from controllers.InitAppController import InitAppController
from controllers.FileDialogController import FileDialogController
from controllers.TableController import TableController

from PyQt6.QtCore import QObject, pyqtSignal
import threading

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class BoletinController(QObject, InitAppController, GenerateNotesController, FileDialogController, TableController):
    """
    Controlador principal para la gestión de la generación de boletines de calificaciones.

    Este controlador maneja la interacción entre la vista principal (`MainView`), 
    la vista de carga (`LoadingView`), y los modelos y controladores necesarios 
    para la creación de carpetas, archivos Excel y PDFs de boletines.

    Hereda de:
        QObject: Clase base para todos los objetos Qt.
        InitAppController: Controlador para la inicialización de la aplicación.
        GenerateNotesController: Controlador para la generación de notas.
        FileDialogController: Controlador para la gestión de diálogos de archivos.
        TableController: Controlador para la gestión de tablas.

    Señales:
        convertion_finished (str): Señal emitida cuando la conversión a PDF ha finalizado.
        convertion_started: Señal emitida cuando comienza la conversión a PDF.
        progress_updated (int, str): Señal emitida para actualizar el progreso de la conversión.
    """

    convertion_finished = pyqtSignal(str)
    convertion_started = pyqtSignal()
    progress_updated = pyqtSignal(int, str) 

    def __init__(self):
        """
        Inicializa el controlador y configura la vista principal y la vista de carga.
        También inicializa la aplicación llamando al método `init_app`.
        """
        super().__init__()
        self.view = MainView()
        self.loading_dialog = LoadingView()
        self.extraction = None
        self.save_path = None
        self.file_path = None
        self.init_app()
        
    def create_folders_boletines(self, params: list, students: list, school_year: str, subjects: list, path: str):
        """
        Crea las carpetas y archivos necesarios para los boletines de calificaciones.

        Args:
            params (list): Parámetros necesarios para la creación de los boletines.
            students (list): Lista de estudiantes para los cuales se generarán los boletines.
            school_year (str): Año escolar para el cual se generarán los boletines.
            subjects (list): Lista de materias que se incluirán en los boletines.
            path (str): Ruta donde se guardarán los archivos generados.

        Este método crea las carpetas necesarias y genera los archivos Excel para cada estudiante.
        Luego, inicia un hilo para la conversión de los archivos Excel a PDF.
        """
        cf = CreateFiles(path, self.convertion_finished, self.progress_updated)
        if cf.create_folders():
            for student in students:
                write = Write(params, student, school_year, subjects, path)
                write.create_excel_boletin()
            
            self.convertion_started.emit()
            thread = threading.Thread(target=cf.create_pdfs_boletin, args=(self.loading_dialog,))
            thread.start()

    def show_loading_dialog(self):
        """
        Muestra el diálogo de carga.

        Este método se utiliza para mostrar la vista de carga mientras se realizan 
        operaciones que pueden tardar un tiempo, como la generación de PDFs.
        """
        self.loading_dialog.show()

    def on_conversion_finished(self, message):
        """
        Maneja la finalización de la conversión a PDF.

        Args:
            message (str): Mensaje que se mostrará al usuario cuando la conversión haya finalizado.

        Este método oculta el diálogo de carga y muestra un mensaje al usuario 
        indicando que la conversión ha finalizado.
        """
        self.loading_dialog.hide()
        self.view.show_message(message, 'information')
    