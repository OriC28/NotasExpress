from models.Extraction.ExtractionNotes import ExtractionNotes
from models.Extraction.ExtractionSubjects import ExtractionSubjects
from models.Extraction.ExtractionSchoolYear import ExtractionSchoolYear
from models.Student.StudentModel import Student, Gradings

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class ExtractionData(ExtractionNotes, ExtractionSubjects, ExtractionSchoolYear):
    """
    Clase encargada de extraer y procesar datos de un archivo Excel para generar información
    sobre estudiantes, materias, calificaciones y año escolar.

    Hereda de:
        ExtractionNotes: Clase para extraer las calificaciones de los estudiantes.
        ExtractionSubjects: Clase para extraer las materias del archivo Excel.
        ExtractionSchoolYear: Clase para extraer el año escolar del archivo Excel.

    Métodos:
        get_specific_data: Extrae datos específicos de una columna en un rango de filas.
        get_data: Obtiene datos de cédulas, apellidos y nombres de los estudiantes.
        filter_students_list: Filtra la lista de estudiantes para eliminar entradas vacías.
        get_students: Crea una lista de objetos `Student` a partir de los datos extraídos.
        save_student_notes: Asigna las calificaciones a cada estudiante en sus respectivas materias.
    """

    def __init__(self, file_path: str, choiced=0):
        """
        Inicializa la clase `ExtractionData`.

        Args:
            file_path (str): Ruta del archivo Excel del cual se extraerán los datos.
            choiced (int, opcional): Índice de la hoja seleccionada en el archivo Excel. Por defecto es 0.
        """
        super().__init__(file_path, choiced)

    def get_specific_data(self, start: int, end: int, column: str):
        """
        Extrae datos específicos de una columna en un rango de filas.

        Args:
            start (int): Fila inicial del rango.
            end (int): Fila final del rango.
            column (str): Letra de la columna de la cual se extraerán los datos.

        Returns:
            list: Lista de datos extraídos de la columna especificada.
        """
        data = []
        for row in range(start, end + 1):
            cell_data = self.sheet_choiced[str(column + str(row))].value
            if cell_data != '**' and cell_data:
                data.append(cell_data)
        return data

    def get_data(self, start: int, end: int):
        """
        Obtiene datos de cédulas, apellidos y nombres de los estudiantes.

        Args:
            start (int): Fila inicial del rango.
            end (int): Fila final del rango.

        Returns:
            list: Lista de listas que contiene cédulas, apellidos y nombres de los estudiantes.
        """
        cedulas = self.get_specific_data(start, end, 'C')
        last_names = self.get_specific_data(start, end, 'D')
        names = self.get_specific_data(start, end, 'E')

        return [cedulas, last_names, names]

    def filter_students_list(self, students):
        """
        Filtra la lista de estudiantes para eliminar entradas vacías.

        Args:
            students (list): Lista de objetos `Student` a filtrar.

        Returns:
            list: Lista de estudiantes filtrada, sin entradas vacías.
        """
        cleansed_students_list = []
        for student in students:
            if student.name and student.last_name and student.cedula:
                cleansed_students_list.append(student)
        return cleansed_students_list

    def get_students(self, start: int, end: int):
        """
        Crea una lista de objetos `Student` a partir de los datos extraídos.

        Args:
            start (int): Fila inicial del rango.
            end (int): Fila final del rango.

        Returns:
            list: Lista de objetos `Student` creados a partir de los datos extraídos.
        """
        students = []
        data = self.get_data(start, end)
        length = len(data[0])
        for i in range(length):
            new_student = Student()
            new_student.cedula = data[0][i]
            new_student.last_name = data[1][i]
            new_student.name = data[2][i]
            students.append(new_student)
        return self.filter_students_list(students)

    def save_student_notes(self, row_subjects: int, table_positions: list, columns: list):
        """
        Asigna las calificaciones a cada estudiante en sus respectivas materias.

        Args:
            row_subjects (int): Fila donde se encuentran los nombres de las materias.
            table_positions (list): Lista con las posiciones de inicio y fin de la tabla de estudiantes.
            columns (list): Lista de columnas donde se encuentran las calificaciones.

        Returns:
            list: Lista de objetos `Student` con las calificaciones asignadas.
        """
        start = table_positions[0]
        end = table_positions[1]
        students = self.get_students(start, end)
        subjects = self.get_subjects(row_subjects)

        for i in range(start, end + 1):
            j = 0
            for subject in subjects:
                students[i - int(start)].subjects_performance[subject.name] = Gradings(subject.name, self.get_notes(i, columns[j]))
                j += 1
        return students