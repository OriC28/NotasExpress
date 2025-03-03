from models.Extraction.ExtractionModel import ExtractionModel
import re

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class ExtractionSchoolYear(ExtractionModel):
    """
    Clase que hereda de `ExtractionModel` para extraer el año escolar de una hoja de cálculo de Excel.

    Attributes:
        file_path (str): Ruta del archivo Excel.
        choiced (int): Índice de la hoja a utilizar. Por defecto es 0.
    """

    def __init__(self, file_path: str, choiced=0):
        """
        Inicializa la clase ExtractionSchoolYear.

        Args:
            file_path (str): Ruta del archivo Excel.
            choiced (int, optional): Índice de la hoja a utilizar. Por defecto es 0.
        """
        super().__init__(file_path, choiced)
    
    def get_school_year(self):
        """
        Extrae el año escolar de una celda específica en la hoja de cálculo.

        Returns:
            str or False: El año escolar en formato 'YYYY-YYYY' si se encuentra, False si no se encuentra.
        """
        date = self.sheet_choiced["B10":"B10"][0][0].value
        row_to_search = date if date is not None else ''  
        year = re.search(r'(\d{4}-\d{4})', row_to_search)
        return year.group().strip() if year is not None else False