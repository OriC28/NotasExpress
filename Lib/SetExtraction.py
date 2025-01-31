# -*- coding: utf-8 -*-
"""
Archivo: SetExtraction.py
Descripción: Módulo que permite obtener los datos de los estudiantes en la hoja seleccionada.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""
# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

# IMPORTANDO MÓDULOS LOCALES
from Lib.Extraction import Extraction


def set_extraction(file_path: str, index_choiced: int):
	"""Permite obtener los datos de los estudiantes en la hoja seleccionada.

	
	Attributes:
	file_path (str): Ruta del archivo de Excel.
	index_choiced (int): Número de la hoja seleccionada.
	
	Returns: 
	list: Lista de estudiantes (clase Student), año escolar (str) y materias (clase Subject).
"""
	extraction = Extraction(file_path, index_choiced)

	# OBTENIENDO LA POSICIÓN DE LAS TABLAS DE NOTAS
	positions = extraction.find_start_end_table(14, 30)
	
	# PARAMETROS A RETORNAR
	school_year = extraction.get_school_year()
	subjects = extraction.get_subjects(13) 

	if not positions or not school_year:
		return False
	# OBTENIENDO LOS DATOS DE LOS ESTUDIANTES
	students = extraction.get_student_data(positions[0], positions[1])
	students = extraction.save_student_notes(13, positions, students)

	#Depurar objetos vacios en total_students
	cleansed_students_list = []
	for student in students:
		if  student.name and student.last_name and student.cedula:
			cleansed_students_list.append(student)

	return cleansed_students_list, school_year, subjects
