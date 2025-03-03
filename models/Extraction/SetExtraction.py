from models.Extraction.ExtractionData import ExtractionData

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

COLUMNS = [
    ["F","G","H","I"],["J","K","L","M"],["N","O","P","Q"],["R","S","T","U"],
    ["V","W","X","Y"],["Z","AA","AB","AC"],["AD","AE","AF","AG"],
    ["AH","AI","AJ","AK"], ['AL','AM','AN','AO']
]

class SetExtraction(ExtractionData):
    """
    Clase que hereda de `ExtractionData` para extraer y organizar todos los datos relevantes de una hoja de cálculo de Excel.

    Attributes:
        file_path (str): Ruta del archivo Excel.
        index_choiced (int): Índice de la hoja a utilizar.
    """

    def __init__(self, file_path:str, index_choiced: int):
        """
        Inicializa la clase SetExtraction.

        Args:
            file_path (str): Ruta del archivo Excel.
            index_choiced (int): Índice de la hoja a utilizar.
        """
        super().__init__(file_path, index_choiced)
    
    def get_all_data(self):
        """
        Extrae y organiza todos los datos relevantes de la hoja de cálculo.

        Returns:
            tuple or False: Una tupla que contiene:
                - Lista de estudiantes con sus notas.
                - El año escolar.
                - Lista de materias (asignaturas).
            Retorna False si no se pueden extraer los datos correctamente.
        """
        table_positions = self.get_start_end_table(14, 30)
        subjects = self.get_subjects(13)
        school_year = self.get_school_year()
        if(table_positions and subjects and school_year):
            students = self.save_student_notes(13, table_positions, COLUMNS)
            return students, school_year, subjects
        return False