from models.Extraction.ExtractionModel import ExtractionModel
from models.Student.StudentModel import Subject

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class ExtractionSubjects(ExtractionModel):
    """
    Clase que hereda de `ExtractionModel` para extraer las materias (asignaturas) de una hoja de cálculo de Excel.

    Attributes:
        file_path (str): Ruta del archivo Excel.
        choiced (int): Índice de la hoja a utilizar. Por defecto es 0.
    """

    def __init__(self, file_path: str, choiced=0):
        """
        Inicializa la clase ExtractionSubjects.

        Args:
            file_path (str): Ruta del archivo Excel.
            choiced (int, optional): Índice de la hoja a utilizar. Por defecto es 0.
        """
        super().__init__(file_path, choiced)
    
    def get_subjects(self, row: int):
        """
        Extrae las materias (asignaturas) de una fila específica en la hoja de cálculo.

        Args:
            row (int): Fila de la cual se extraerán las materias.

        Returns:
            list: Lista de objetos `Subject` que representan las materias encontradas.
        """
        subjects = []
        for i in self.sheet_choiced[row]:
            if i.value is not None and i.value!="Promedios":
                subjects.append(Subject(i.value.strip()))
        return subjects