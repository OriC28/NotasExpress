from models.PDF.GetPDF import run_export_in_background
from PyQt6.QtCore import pyqtSignal
import threading
import os

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class CreateFiles:
    """
    Clase para crear carpetas y generar archivos PDF a partir de archivos Excel.

    Attributes:
        path (str): Ruta base donde se crearán las carpetas y se guardarán los archivos.
        convertion_finished (pyqtSignal): Señal emitida cuando se completa la conversión.
        progress_updated (pyqtSignal): Señal emitida para actualizar el progreso de la conversión.
    """

    def __init__(self, path:str, convertion_finished: pyqtSignal, progress_updated: pyqtSignal):
        """
        Inicializa la clase CreateFiles.

        Args:
            path (str): Ruta base donde se crearán las carpetas y se guardarán los archivos.
            convertion_finished (pyqtSignal): Señal emitida cuando se completa la conversión.
            progress_updated (pyqtSignal): Señal emitida para actualizar el progreso de la conversión.
        """
        self.path = path
        self.convertion_finished = convertion_finished
        self.progress_updated = progress_updated

    def create_folders(self):
        """
        Crea las carpetas necesarias para almacenar los archivos PDF y Excel.

        Returns:
            bool: True si las carpetas se crearon correctamente, False si ya existían.

        Raises:
            Exception: Si las carpetas ya existen en el directorio especificado.
        """
        created = False
        if not os.path.exists(self.path):
            os.mkdir(self.path)
            os.mkdir(os.path.join(self.path, "PDF"))
            os.mkdir(os.path.join(self.path, "EXCEL"))  
            created = True
        else:
            raise Exception('Ya existen boletines de esta sección en el directorio seleccionado. Por favor, cambie el directorio o elimine los archivos conflictivos.')
        return created

    def progress_loading(self, current_step, total_steps, file):
        """
        Calcula y emite el progreso actual de la conversión.

        Args:
            current_step (int): Paso actual del proceso.
            total_steps (int): Total de pasos del proceso.
            file (str): Nombre del archivo actual que se está procesando.
        """
        progress = int((current_step / total_steps) * 100)
        self.progress_updated.emit(progress, file)  

    def create_pdfs_boletin(self, loading_dialog):
        """
        Convierte todos los archivos Excel en la carpeta 'EXCEL' a PDF y los guarda en la carpeta 'PDF'.

        Verifica si la bandera de cancelación está establecida para terminar el proceso de 
        generación de boletines en formato PDF, en caso contrario mostrará un mensaje de éxito.
        Además, emite señales para actualizar el progreso y notificar la finalización del proceso.
        
        Args:
            loading_dialog (LoadingView): Ventana de carga.
        """
        files = [i for i in os.listdir(os.path.join(self.path, 'EXCEL')) if i.endswith(".xlsx")]
        total_files = len(files)

        for i, file in enumerate(files, start=1):
            
            if loading_dialog.cancel_flag.is_set():
                break
            output_file = os.path.join(self.path, 'PDF', f"{file.split('.')[0]}.pdf")
            input_file = os.path.join(self.path, 'EXCEL', file)

            input_file = input_file.replace("/", '\\')
            output_file = output_file.replace('/', '\\')

            self.progress_loading(i, total_files, output_file) 
            run_export_in_background(input_file, output_file)
        
        if not loading_dialog.cancel_flag.is_set():
            self.convertion_finished.emit('Se ha finalizado la generación de boletines con éxito.')
        else:
            loading_dialog.cancel_flag = threading.Event()
            self.convertion_finished.emit('La generación de boletines fue cancelada.')