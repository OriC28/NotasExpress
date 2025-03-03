from PyQt6.QtWidgets import QDialog, QMessageBox
from PyQt6 import uic
from PyQt6.QtCore import Qt
import threading

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class LoadingView(QDialog):
    """
    Clase que representa una ventana de diálogo para mostrar el progreso de una operación.

    Esta ventana se mantiene siempre visible encima de otras ventanas y muestra una barra de progreso
    junto con el nombre del archivo que se está procesando.

    Attributes:
        Ninguno.
    """

    def __init__(self):
        """
        Inicializa la ventana de carga.

        Carga la interfaz de usuario desde el archivo 'GUI/Loading.ui' y configura la ventana
        para que permanezca siempre visible encima de otras ventanas.
        """
        super().__init__()
        uic.loadUi('GUI/Loading.ui', self)
        self.message = QMessageBox()
        self.cancel_flag = threading.Event()
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)

    def update_progress(self, value, file):
        """
        Actualiza el progreso y el nombre del archivo que se está procesando.

        Args:
            value (int): Valor del progreso (0-100).
            file (str): Nombre del archivo que se está procesando.
        """
        self.files.setText(f'Generando {file}...')
        self.progressBar.setValue(value)
    """
    Confirma el cierre de la ventana de carga.

    Este método se ejecuta cuando el usuario intenta cerrar la ventana de carga. 
    Muestra un diálogo de confirmación para preguntar si el usuario está seguro 
    de detener la generación de boletines. Si el usuario confirma,se establece 
    una bandera de cancelación y se acepta el evento de cierre. De lo contrario, 
    se ignora el evento.

    Args:
        event (QCloseEvent): Evento de cierre que se activa cuando el usuario intenta cerrar la ventana.

    Returns:
        None
    """
    def closeEvent(self, event):
        answer = self.message.question(
            self, 'Detener generación', "¿Está seguro que desea detener la generación de boletines?",
            QMessageBox.StandardButton.Yes| QMessageBox.StandardButton.No
        )
        if answer == 16384:
            event.accept()
            self.cancel_flag.set()
        else:
            event.ignore()
