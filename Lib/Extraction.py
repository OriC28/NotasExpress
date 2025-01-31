# -*- coding: utf-8 -*-
"""
Archivo: Extration.py
Descripción: Módulo que permite extraer los datos de un archivo excel.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

# IMPORTANDO LIBRERIAS NECESARIAS
from openpyxl import load_workbook
import re

# IMPORTANDO MÓDULOS LOCALES
from Lib.students import Student, Gradings, Subject

# COLUMNAS DE LAS NOTAS
COLUMNS = [["F","G","H","I"],["J","K","L","M"],["N","O","P","Q"],
		   ["R","S","T","U"],["V","W","X","Y"],["Z","AA","AB","AC"],
		   ["AD","AE","AF","AG"],["AH","AI","AJ","AK"], ["AL","AM","AN","AO"],
		   ["AP","AQ","AR","AS"],["AT","AU","AV","AW"]]

class Extraction:
	"""
	Clase que permite extraer los datos de un archivo Excel.
	

	Attributes:
	file_path (str): Ruta del archivo de Excel.
	choiced (int): Número de la hoja seleccionada.
	workbook (Workbook): Archivo de Excel.
	sheets (list): Lista de hojas del archivo de Excel.
	sheet_choiced (Worksheet): Hoja seleccionada.
	"""
	def __init__(self, file_path: str, choiced=0):
		"""Inicialización de la clase Extraction.
		
		
		Attributes:
		file_path (str) : Ruta del archivo de Excel.
		choiced (int): Número de la hoja seleccionada."""
		self.file_path = file_path
		self.choiced = choiced
		self.workbook = load_workbook(self.file_path, data_only=True)
		self.sheets = self.workbook.sheetnames
		self.sheet_choiced = self.workbook[self.sheets[choiced]]

	def find_start_end_table(self, aprox_start: int, aprox_end: int):
		start = None
		end = None
		"""Encuentra el inicio y final de la tabla de notas de los 
		estudiantes en la hoja seleccionada sin el encabezado.
	
	
		Attributes:
		aprox_start (int) : Nro aproximado de la celda inicial de la tabla.
		aprox_end (int) : Nro aproximado de la última celda de la tabla.

		Returns:
		list [start, end] : Arreglo con la posición de la celda inicial y la posición de la celda final
		"""
		# OBTENIENDO EL INICIO DE LA TABLA SIN ENCABEZADO
		for n in range(aprox_start, aprox_end):
			if re.findall(r'^V-|^CE-', str(self.sheet_choiced['C' + str(n)].value)):
				start = n
				break
			
		if start is None:
			return False
		
		# OBTENIENDO EL FINAL DE LA TABLA SIN ENCABEZADO
		for n in range(start, aprox_end):
			if self.sheet_choiced['C' + str(n)].value is None:
				end = n-1
				break

		return [start, end]
		
	def get_student_data(self, start: int, end: int):
		"""Permite obtener los datos de los estudiantes en la hoja seleccionada.


		Attributes:
		start (int): Posición inicial de la tabla de estudiantes.
		end (int): Posición final de la tabla de estudiantes.
		
		Returns: 
		list: Lista de estudiantes.
		"""
		students_list = []
		for row in range(start, end+1):
			new_student = Student()
			# OBTENIENDO LOS DATOS DE LOS ESTUDIANTES
			for column in 'CDE':
				# OBTENIENDO EL VALOR DE LA CELDA EN LA HOJA DE EXCEL 
				cell_data = self.sheet_choiced[str(column+str(row))].value
				if cell_data!='**' and cell_data:
					# ASIGNANDO LOS DATOS A LOS ATRIBUTOS DEL ESTUDIANTE
					match column:
						case'C':
							new_student.cedula = cell_data
						case 'D':
							new_student.last_name = cell_data
						case'E':
							new_student.name = cell_data
			# AGREGANDO LOS ESTUDIANTES A LA LISTA
			students_list.append(new_student)
		return students_list
	
	def get_subjects(self, row: int):
		"""
		Obtiene las asignaturas de la hoja seleccionada.

		Attributes:
		row (int): Fila donde se encuentran las asignaturas.
		
		Returns: 
		list: Lista de asignaturas (clase Subject).
		"""
		subjects = []
		for i in self.sheet_choiced[row]:
			# OBTENIENDO EL NOMBRE DE LA ASIGNATURA
			if i.value is not None and i.value!="Promedios":
				# AGREGANDO LA ASIGNATURA A LA LISTA
				subjects.append(Subject(i.value))
		return subjects
	
	def get_notes(self, row: int, block: list):
		"""Obtiene las notas de los estudiantes en la hoja seleccionada.
	

		Attributes:
		row (int): Fila donde se encuentran las notas.
		block (list): Bloque de columnas donde se encuentran las notas.
	
		Returns: 
		list: Lista de notas.
		"""
		notes = []
		for letter in block:
			# OBTENIENDO LAS NOTAS DE LOS ESTUDIANTES
			if self.sheet_choiced[letter + str(row)].value is None:
				note = self.sheet_choiced[letter + str(row)].value  = '**'
			else:
				note = round(float(self.sheet_choiced[letter + str(row)].value), 2)
			# AGREGANDO LAS NOTAS A LA LISTA
			notes.append(note)
		return notes

	def save_student_notes(self, row_subjects: int, table_positions: list, students: list):
		"""Guarda las notas de cada estudiante de manera individual en los objetos de la clase Student.


		Attributes:
		row_subjects (int): Fila donde se encuentran las asignaturas.
		table_position (list): Posición de inicio y final de la tabla.
		students (list): Lista de estudiantes.
		
		Returns: 
		list: Lista de estudiantes con las notas.
		"""
		i = 0
		# OBTENIENDO LAS POSICIONES DE INICIO Y FINAL DE LA TABLA
		start = table_positions[0]
		end = table_positions[1]
		# OBTENIENDO LAS ASIGNATURAS
		subjects = self.get_subjects(row_subjects)
		for i in range (start, end+1):
			j = 0
			for subject in subjects:
				# ASIGNANDO LAS NOTAS A CADA ESTUDIANTE
				students[i-int(start)].subjects_performance[subject.name] = Gradings(subject.name, self.get_notes(i, COLUMNS[j]))
				j += 1
		return students 
	
	def get_school_year(self):
		"""
		Obtiene el año escolar de la hoja seleccionada.

		Returns:
		str: Año escolar.
		"""
		# FILA EN LA CUAL BUSCAR EL AÑO ESCOLAR
		date = self.sheet_choiced["B10":"B10"][0][0].value
		row_to_search = date if date is not None else ''
		# RECUPERANDO ÚNICAMENTE EL AÑO ESCOLAR
		year = re.search(r'(\d{4}-\d{4})', row_to_search)
		if year is not None:
			return year.group()
		return False