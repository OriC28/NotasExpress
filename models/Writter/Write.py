from openpyxl import load_workbook
import os

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class Write:
    """
    Clase para escribir datos en un archivo Excel (boletín) y guardarlo.

    Attributes:
        params (list): Lista de parámetros adicionales (hoja, mención, profesor guía, fecha).
        student (Student): Objeto que representa al estudiante.
        school_year (str): Año escolar.
        subjects (list): Lista de materias (asignaturas).
        path (str): Ruta donde se guardará el archivo Excel.
    """

    def __init__(self, params:list, student, school_year: str, subjects: list, path: str):
        """
        Inicializa la clase Write.

        Args:
            params (list): Lista de parámetros adicionales (hoja, mención, profesor guía, fecha).
            student (Student): Objeto que representa al estudiante.
            school_year (str): Año escolar.
            subjects (list): Lista de materias (asignaturas).
            path (str): Ruta donde se guardará el archivo Excel.
        """
        self.boletin = load_workbook('resources/BOLETIN.xlsx', data_only=True)
        self.sheet = self.boletin.active
        self.path = path
        self.student = student
        self.school_year = school_year
        self.subjects = subjects
        self.sheet_choiced = params[0]
        self.mention = params[1]
        self.guide_teacher = params[2]
        self.date = params[3]
        
    def set_moment(self, column:str, row: int, note: str):
        """
        Establece la nota de un momento específico en una celda del boletín.

        Args:
            column (str): Columna de la celda.
            row (int): Fila de la celda.
            note (str): Nota a establecer.
        """
        self.sheet[column + str(row)].value = note
    
    def create_excel_boletin(self):
        """
        Crea el boletín en formato Excel con los datos del estudiante, materias y notas.
        """
        if self.student.name != str:
            # AÑO ESCOLAR
            self.sheet['C4'].value += self.school_year
            # NOMBRE COMPLETO
            self.sheet['A5'].value += f"{self.student.name} {self.student.last_name}"
            # CÉDULA DEL ESTUDIANTE
            self.sheet['I5'].value += self.student.cedula
            # MENCIÓN 
            self.sheet['A6'].value += self.mention
            # AÑO-SECCIÓN
            self.sheet['D6'].value += self.sheet_choiced
            # PROFESOR GUÍA
            self.sheet['F6'].value += self.guide_teacher
            # FECHA DE ENTREGA
            self.sheet['J6'].value += self.date

            self.set_notes_in_boletin()

    def verify_notes(self, student_data, moment):
        """
        Verifica y retorna la nota de un momento específico.

        Args:
            student_data (dict): Datos del estudiante.
            moment (int): Momento (0, 1, 2, 3).

        Returns:
            float: Nota del momento especificado.
        """
        def_note = 0
        if student_data.moment_grades[moment] != "**":
            def_note += student_data.moment_grades[moment]
        return def_note

    def set_definives_notes(self, def_first_moment, def_second_moment, def_third_moment, def_general):
        """
        Establece las notas definitivas en el boletín.

        Args:
            def_first_moment (float): Nota definitiva del primer momento.
            def_second_moment (float): Nota definitiva del segundo momento.
            def_third_moment (float): Nota definitiva del tercer momento.
            def_general (float): Nota definitiva general.
        """
        self.sheet['H18'].value = str(round(def_first_moment / len(self.subjects), 2))
        self.sheet['I18'].value = str(round(def_second_moment / len(self.subjects), 2))
        self.sheet['J18'].value = str(round(def_third_moment / len(self.subjects), 2))
        self.sheet['K18'].value = str(round(def_general / len(self.subjects), 2))
        self.sheet['K20'].value = str(round(def_general / len(self.subjects), 2))
    
    def set_notes_in_boletin(self):
        """
        Establece las notas de todas las materias en el boletín.
        """
        def_general = 0
        def_first_moment = 0
        def_second_moment = 0
        def_third_moment = 0
        row = 9
        for subject in self.subjects:
            if row > 19:
                break
            else:
                # MATERIAS
                self.sheet['A' + str(row)].value = subject.name.upper()
                student_data = self.student.subjects_performance[subject.name]

                self.set_moment('H', row, str(student_data.moment_grades[0]))
                self.set_moment('I', row, str(student_data.moment_grades[1]))
                self.set_moment('J', row, str(student_data.moment_grades[2]))
                self.set_moment('K', row, str(student_data.moment_grades[3]))

                def_first_moment += self.verify_notes(student_data, 0)
                def_second_moment += self.verify_notes(student_data, 1)
                def_third_moment += self.verify_notes(student_data, 2)
                def_general += self.verify_notes(student_data, 3)
                
            row += 1
        self.set_definives_notes(def_first_moment, def_second_moment, def_third_moment, def_general)
        self.save_boletin(self.student.cedula)

    def save_boletin(self, cedula: str):
        """
        Guarda el boletín en formato Excel.

        Args:
            cedula (str): Cédula del estudiante, usada como nombre del archivo.
        """
        self.boletin.save(os.path.join(self.path, 'EXCEL', f'{cedula}.xlsx'))
        self.boletin.close()