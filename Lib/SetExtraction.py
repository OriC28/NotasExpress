# -*- coding: utf-8 -*-
"""
Archivo: SetExtraction.py
Descripción: Módulo que permite obtener los datos de los estudiantes en la hoja seleccionada.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""
# IMPORTANDO MÓDULOS LOCALES
from Lib.Extraction import Extraction

"""
Permite obtener los datos de los estudiantes en la hoja seleccionada.

@param file_path (str): Ruta del archivo de Excel.
@param index_choiced (int): Número de la hoja seleccionada.
returns list: Lista de estudiantes, año escolar y materias.
"""
def set_extraction(file_path: str, index_choiced: int):
	extraction = Extraction(file_path, index_choiced)

	# OBTENIENDO LA POSICIÓN DE LAS TABLAS DE NOTAS
	first_table_position = extraction.find_start_end_table(14, 30)
	second_table_position = extraction.find_start_end_table(first_table_position[1]+1, 60)  

	# OBTENIENDO LOS DATOS DE LOS ESTUDIANTES
	students = extraction.get_student_data(first_table_position[0], first_table_position[1])
	students = extraction.save_student_notes(14, first_table_position, students)

	# PARAMENTROS A RETORNAR
	total_students = extraction.save_student_notes(second_table_position[0]-2, second_table_position, students)
	school_year = extraction.get_school_year()
	subjects = extraction.get_subjects(14) + extraction.get_subjects(second_table_position[0]-2)
	
	return total_students, school_year, subjects