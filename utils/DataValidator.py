import re

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class DataValidator:

    @staticmethod
    def validate_file_input(extraction, view):
        if extraction:
            if extraction.get_all_data() == False:
                view.clean_all_in_gui()
                raise Exception("Archivo seleccionado inválido.")


    """
    Clase que proporciona métodos estáticos para validar campos de entrada.

    Esta clase se utiliza para verificar que los campos de entrada no estén vacíos y cumplan con ciertos patrones de formato.
    """

    @staticmethod
    def validate_fields(inputs):
        """
        Valida que los campos de entrada no estén vacíos y cumplan con los patrones de formato especificados.

        Args:
            inputs (dict): Diccionario que contiene los campos a validar. Las claves son los nombres de los campos y los valores son los datos ingresados.

        Raises:
            Exception: Si algún campo está vacío o no cumple con el formato requerido.
        """
        for input in inputs:
            if inputs[input] == "":
                raise Exception(f'El campo de {input} debe ser llenado para la generación de los archivos.')
        DataValidator.preg_match(inputs)

    @staticmethod
    def preg_match(inputs):
        """
        Valida que los campos de entrada cumplan con los patrones de formato especificados.

        Args:
            inputs (dict): Diccionario que contiene los campos a validar. Las claves son los nombres de los campos y los valores son los datos ingresados.

        Raises:
            Exception: Si el campo 'Profesor guía' contiene caracteres no permitidos o si el campo 'Fecha' no sigue el formato "día/mes/año".
        """
        if not re.match("^[a-zA-ZáéíóúÁÉÍÓÚüÜñÑ'\\s]+$", inputs['Profesor guía']):
            raise Exception('Solo se pueden colocar letras como carácteres en el campo de profesor guía. Por favor, cambialos para continuar')
        elif not re.match("^([1-9]|[12][0-9]|3[01])\\/([1-9]|0[1-9]|1[0-2])\\/(19|20)\\d{2}$", inputs['Fecha']):
            raise Exception('Asegúrate de que la fecha está bien escrita, debe seguir el formato: "día/mes/año" .')