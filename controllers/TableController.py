from models.Extraction.SetExtraction import SetExtraction

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class TableController:
    """
    Controlador para gestionar la interacción con la tabla de la interfaz gráfica.

    Este controlador se encarga de:
    1. Obtener el índice seleccionado en el combo box de año/sección.
    2. Llenar la tabla con los datos de los estudiantes extraídos del archivo Excel.
    3. Calcular el promedio de las calificaciones de cada estudiante.

    Métodos:
        get_index_choiced: Obtiene el índice seleccionado en el combo box.
        fill_table: Llena la tabla con los datos de los estudiantes.
        calculate_average_score: Calcula el promedio de las calificaciones de un estudiante.
    """

    def get_index_choiced(self):
        """
        Obtiene el índice seleccionado en el combo box de año/sección.

        Si no hay ningún índice seleccionado, se devuelve 0 por defecto.

        Returns:
            int: Índice seleccionado en el combo box.
        """
        index_choiced = self.view.CbYearSection.currentIndex()
        if index_choiced < 0:
            index_choiced = 0
        return index_choiced

    def fill_table(self):
        """
        Llena la tabla con los datos de los estudiantes extraídos del archivo Excel.

        Este método realiza las siguientes acciones:
        1. Obtiene el índice seleccionado en el combo box.
        2. Extrae los datos del archivo Excel utilizando el índice seleccionado.
        3. Si la extracción es exitosa, obtiene los datos de los estudiantes.
        4. Configura el número de filas en la tabla.
        5. Establece las menciones en la interfaz gráfica.
        6. Llena la tabla con los datos de los estudiantes y sus promedios.

        Excepciones:
            Muestra un mensaje de advertencia si ocurre un error durante el proceso.
        """
        try:
            index_choiced = self.get_index_choiced()
            self.extraction = SetExtraction(self.file_path, index_choiced)
            if self.extraction is not None:
                data = self.extraction.get_all_data()
                if data != False:
                    total_students = data[0]
                    row_count = len(total_students) if len(total_students) >= 8 else 8
                    self.view.set_rows_table(row_count)
                    self.view.set_mentions(['TRANSPORTE ACUÁTICO', 'METALMECÁNICA'])
                    self.view.fill_table(total_students, self.calculate_average_score)
            else:
                self.view.clean_all_in_gui()
        except Exception as e:
            self.view.show_message(f'Ha ocurrido un error inesperado al cargar la tabla: {e}', 'warning')

    def calculate_average_score(self, student):
        """
        Calcula el promedio de las calificaciones de un estudiante.

        Este método suma las calificaciones del cuarto momento (índice 3) de todas las materias
        y las divide por el número de materias para obtener el promedio.

        Args:
            student: Objeto que representa a un estudiante con sus calificaciones.

        Returns:
            float: Promedio de las calificaciones del estudiante, redondeado a 2 decimales.
        """
        total_notes = sum(
            subject_performance.moment_grades[3]
            for subject_performance in student.subjects_performance.values() if student.subjects_performance.values() != "**"
        )
        return round(total_notes / len(student.subjects_performance), 2)