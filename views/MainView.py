from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox, QTableWidget, QTableWidgetItem, QFileDialog
from PyQt6 import uic, QtGui, QtCore
import os

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class MainView(QMainWindow):
    """
    Clase que representa la ventana principal de la aplicación.

    Esta clase maneja la interfaz gráfica de usuario (GUI) y proporciona métodos para interactuar
    con los widgets, mostrar mensajes, abrir diálogos de archivos y llenar una tabla con datos.

    Attributes:
        message (QMessageBox): Objeto para mostrar mensajes al usuario.
        table (QTableWidget): Objeto para manejar la tabla de la interfaz.
    """

    def __init__(self):
        """
        Inicializa la ventana principal.

        Carga la interfaz de usuario desde el archivo 'gui/GUI.ui', centra la ventana en la pantalla
        y configura el ícono de la aplicación.
        """
        super().__init__()
        uic.loadUi('gui/GUI.ui', self)
        self.message = QMessageBox()
        self.table = QTableWidget
        self.center()

    def center(self):
        """
        Centra la ventana principal en la pantalla.
        """
        screen_geometry = QApplication.primaryScreen().geometry()

        window_geometry = self.frameGeometry()

        x = (screen_geometry.width() - window_geometry.width()) // 2
        y = (screen_geometry.height() - window_geometry.height()) // 2

        self.move(x, y)

    def set_logo(self):
        """
        Configura el logo de la aplicación en la interfaz gráfica.

        Carga la imagen del logo desde 'resources/LOGO.png' y la muestra en el widget correspondiente.
        También establece el ícono de la ventana.
        """
        imagen = QtGui.QPixmap("resources/LOGO.png")
        self.logo.setScaledContents(True)
        self.logo.resize(imagen.width(), imagen.height())
        self.logo.setPixmap(imagen)
        self.setWindowIcon(QtGui.QIcon('resources/icons/iconProgram.ico'))
    
    def open_file_dialog_to_save(self):
        """
        Abre un diálogo para seleccionar un directorio donde guardar archivos.

        Returns:
            str: Ruta del directorio seleccionado.
        """
        dir_path = QFileDialog.getExistingDirectory(
            parent=self, caption="Select directory",
            directory=os.path.expanduser('~'),        
            options=QFileDialog.Option.DontUseNativeDialog    
        )
        return dir_path
    
    def open_file_dialog_to_file(self):
        """
        Abre un diálogo para seleccionar un archivo Excel.

        Returns:
            str: Ruta del archivo seleccionado.
        """
        file_path, filter = QFileDialog.getOpenFileName(self, 'Open file', '', 'Excel files (*.xlsx)')
        return file_path
    
    def clean_all_before_open_file_dialog(self):
        """
        Limpia los widgets relacionados con la selección de archivos antes de abrir un diálogo.
        """
        self.CbYearSection.clear()
        self.CbYearSection.setEnabled(False)
        self.pathFile.setText('')
    
    def clean_all_in_gui(self):
        """
        Limpia todos los widgets de la interfaz gráfica.
        """
        self.CbYearSection.clear()
        self.CbMention.clear()
        self.CbYearSection.setEnabled(False)
        self.CbMention.setEnabled(False)
        self.Table.clearContents()

    def get_data_inputs(self):
        """
        Obtiene los datos ingresados en los campos de texto de la interfaz.

        Returns:
            dict: Diccionario con los datos ingresados (profesor guía y fecha).
        """
        guide_teacher = self.GuideTeacherEntry.text()
        date = self.DateEntry.text()
        return {"Profesor guía": guide_teacher, "Fecha": date}
    
    def get_all_data_selected(self):
        """
        Obtiene todos los datos seleccionados o ingresados en la interfaz.

        Returns:
            list: Lista con los datos seleccionados (año/sección, mención, profesor guía y fecha).
        """
        sheet_choiced_name = self.CbYearSection.currentText()
        mention = self.CbMention.currentText()
        guide_teacher = self.GuideTeacherEntry.text()
        date = self.DateEntry.text()
        return [sheet_choiced_name, mention, guide_teacher, date]
    
    def set_sheets(self, sheets: list):
        """
        Configura las opciones del combo box de año/sección.

        Args:
            sheets (list): Lista de opciones para el combo box.
        """
        self.CbYearSection.addItems(sheets)

    def set_mentions(self, mentions: list):
        """
        Configura las opciones del combo box de mención.

        Args:
            mentions (list): Lista de opciones para el combo box.
        """
        self.CbMention.setEnabled(True)
        if self.CbMention.count() == 0:
            self.CbMention.addItems(mentions)
    
    def set_rows_table(self, row):
        """
        Configura el número de filas de la tabla.

        Args:
            row (int): Número de filas a establecer.
        """
        self.Table.clearContents()
        self.Table.setRowCount(row)

    def activate_widgets(self):
        """
        Habilita los widgets de la interfaz gráfica.
        """
        self.CbYearSection.setEnabled(True)
        self.SaveButton.setEnabled(True)

    def fill_table(self, students: list, calculate_average_score: callable):
        """
        Llena la tabla con los datos de los estudiantes.

        Args:
            students (list): Lista de objetos `Student` que representan a los estudiantes.
            calculate_average_score (callable): Función para calcular el promedio de notas de un estudiante.
        """
        for index, student in enumerate(students):
            if student.name:
                self.add_student_to_table(index, student, calculate_average_score(student))

    def add_student_to_table(self, index, student, average_score):
        """
        Agrega un estudiante a la tabla.

        Args:
            index (int): Índice de la fila en la tabla.
            student (Student): Objeto que representa al estudiante.
            average_score (float): Promedio de notas del estudiante.
        """
        student_data = [
            str(index + 1),
            str(student.cedula),
            str(student.name),
            str(student.last_name),
            str(average_score)
        ]
        
        for column, data in enumerate(student_data):
            self.Table.setItem(index, column, QTableWidgetItem(data))

    def show_message(self, message: str, type: str):
        """
        Muestra un mensaje al usuario.

        Args:
            message (str): Mensaje a mostrar.
            type (str): Tipo de mensaje ("warning", "information" o "question").
        """
        if type == "warning":
            self.message.warning(self, type, message, QMessageBox.StandardButton.Ok)
        elif type == "information":
            self.message.information(self, type, message, QMessageBox.StandardButton.Ok)
        elif type == 'question':
            return self.show_dialog(message, type)
    
    def show_dialog(self, message: str, type: str):
        """
        Muestra un diálogo de pregunta al usuario.

        Args:
            message (str): Mensaje a mostrar.
            type (str): Tipo de diálogo ("question").

        Returns:
            bool: True si el usuario selecciona "Sí", False si selecciona "No".
        """
        answer = self.message.question(self, type, message, QMessageBox.StandardButton.Yes
                                        | QMessageBox.StandardButton.No)
        return True if answer == 16384 else False

    """
    Confirma el cierre de la ventana de principal.

    Este método se ejecuta cuando el usuario intenta cerrar la ventana principal. 
    Muestra un diálogo de confirmación para preguntar si el usuario está seguro 
    salir del programa. Si el usuario confirma, se acepta el evento de cierre. 
    De lo contrario, se ignora el evento.

    Args:
        event (QCloseEvent): Evento de cierre que se activa cuando el usuario intenta cerrar la ventana.

    Returns:
        None
    """
    def closeEvent(self, event):
        answer = self.show_dialog('¿Está seguro que desea cerrar el programa?', 'Cerrando programa')
        if answer:
            event.accept()
            import sys
            sys.exit(0)
        else:
            event.ignore()
