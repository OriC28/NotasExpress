from models.Extraction.ExtractionModel import ExtractionModel

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class ExtractionNotes(ExtractionModel):
    """
    Clase que hereda de `ExtractionModel` para extraer notas de una hoja de cálculo de Excel.

    Attributes:
        file_path (str): Ruta del archivo Excel.
        choiced (int): Índice de la hoja a utilizar. Por defecto es 0.
    """

    def __init__(self, file_path: str, choiced=0):
        """
        Inicializa la clase ExtractionNotes.

        Args:
            file_path (str): Ruta del archivo Excel.
            choiced (int, optional): Índice de la hoja a utilizar. Por defecto es 0.
        """
        super().__init__(file_path, choiced)

    def get_notes(self, row: int, block: list):
        """
        Obtiene las notas de una fila específica para un conjunto de celdas definido por `block`.

        Args:
            row (int): Fila de la cual se extraerán las notas.
            block (list): Lista de letras que representan las columnas de las celdas a extraer.

        Returns:
            list: Lista de notas redondeadas a 2 decimales o '**' si la celda está vacía o contiene '**'.
        """
        notes = []
        for letter in block:
            cell = self.sheet_choiced[letter + str(row)].value
            if cell is not None and cell !='**':
                note = round(float(self.sheet_choiced[letter + str(row)].value), 2)
            else:
                note = self.sheet_choiced[letter + str(row)].value  = '**'
            notes.append(note)
        return notes