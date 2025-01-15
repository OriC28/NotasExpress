# -*- coding: utf-8 -*-
"""
Archivo: Extration.py
Descripción: Módulo que permite extraer los datos de un archivo excel.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""
# IMPORTANDO LIBRERIAS NECESARIAS
from openpyxl import load_workbook
import re

# IMPORTANDO MÓDULOS LOCALES
from Lib.students import Student, Gradings, Subject

# COLUMNAS DE LAS NOTAS
COLUMNS = [["F","G","H","I"],["J","K","L","M"],["N","O","P","Q"],["R","S","T","U"],["V","W","X","Y"]]

"""
Clase que permite extraer los datos de un archivo excel.
	
@attr file_path (str): Ruta del archivo de Excel.
@attr choiced (int): Número de la hoja seleccionada.
@attr workbook (Workbook): Archivo de Excel.
@attr sheets (list): Lista de hojas del archivo de Excel.
@attr sheet_choiced (Worksheet): Hoja seleccionada.
"""
class Extraction:
	def __init__(self, file_path: str, choiced=0):
		self.file_path = file_path
		self.choiced = choiced
		self.workbook = load_workbook(self.file_path, data_only=True)
		self.sheets = self.workbook.sheetnames
		self.sheet_choiced = self.workbook[self.sheets[choiced]]

	"""
	Permite encontrar el inicio y final de la tabla de notas de los 
	estudiantes en la hoja seleccionada sin el encabezado.
		
	@param aprox_start (int): Aproximación del inicio de la tabla.
	@param aprox_end (int): Aproximación del final de la tabla.
	returns list: Posición de inicio y final de la tabla.
	"""
	def find_start_end_table(self, aprox_start: int, aprox_end: int):
		# OBTENIENDO EL INICIO DE LA TABLA SIN ENCABEZADO
		for n in range(aprox_start, aprox_end):
			if re.findall(r'^V-|^CE-', str(self.sheet_choiced['C' + str(n)].value)):
				start = n
				break
		# OBTENIENDO EL FINAL DE LA TABLA SIN ENCABEZADO
		for n in range(start, aprox_end):
			if self.sheet_choiced['C' + str(n)].value is None:
				end = n-1
				break
		return [start,end]
	
	"""
	Permite obtener los datos de los estudiantes en la hoja seleccionada.

	@param start (int): Posición inicial de la tabla de estudiantes.
	@param end (int): Posición final de la tabla de estudiantes.
	returns list: Lista de estudiantes.
	"""
	def get_student_data(self, start: int, end: int):
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
	"""
	Obtiene las asignaturas de la hoja seleccionada.

	@param row (int): Fila donde se encuentran las asignaturas.
	returns list: Lista de asignaturas.
	"""
	def get_subjects(self, row: int):
		subjects = []
		for i in self.sheet_choiced[row]:
			# OBTENIENDO EL NOMBRE DE LA ASIGNATURA
			if i.value is not None and i.value!="Promedios":
				# AGREGANDO LA ASIGNATURA A LA LISTA
				subjects.append(Subject(i.value))
		return subjects
	"""
	Obtniene las notas de los estudiantes en la hoja seleccionada.

	@param row (int): Fila donde se encuentran las notas.
	@param block (list): Bloque de columnas donde se encuentran las notas.
	returns list: Lista de notas.
	"""
	def get_notes(self, row: int, block: list):
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

	"""
	Guarda las notas de cada estudiante de manera individual en los objetos de la clase Student.

	@param row_subjects (int): Fila donde se encuentran las asignaturas.
	@param table_position (list): Posición de inicio y final de la tabla.
	@param students (list): Lista de estudiantes.
	returns list: Lista de estudiantes con las notas.
	"""
	def save_student_notes(self, row_subjects: int, table_positions: list, students: list):
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
	"""
	Obtiene el año escolar de la hoja seleccionada.

	returns str: Año escolar.
	"""
	def get_school_year(self):
		# FILA EN LA CUAL BUSCAR EL AÑO ESCOLAR
		row_to_search = self.sheet_choiced["B11":"B11"][0][0].value
		# RECUPERANDO ÚNICAMENTE EL AÑO ESCOLAR
		year = re.search(r'(\d{4}-\d{4})', row_to_search)
		return year.group()