
"""
Archivo: WriteFile.py
Descripción: Este archivo contiene las funciones necesarias para crear los archivos de los boletines de los estudiantes.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""
# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

# IMPORTANDO LIBRERÍAS NECESARIAS
from openpyxl import load_workbook
import os

# IMPORTANDO MÓDULOS LOCALES
from Lib.GetPDF import run_export_in_background


def create_folders(path: str):
	"""Crea los directorios necesarios para almacenar los archivos de los boletines

	Si existe un directorio homónimo, entonces arrojará una excepción con dicha advertencia.

	Attributes:
	path (str) : Ruta del directorio principal.

	Returns: 
	created (bool) : True si se crearon los directorios, False en caso contrario."""
	created = False
	if not os.path.exists(path):
		# CREA EL DIRECTORIO PRINCIPAL 
		os.mkdir(path)
		# CREA UN SUBDIRECTORIO LLAMADO "PDF" 
		os.mkdir(os.path.join(path, "PDF"))
		# CREA UN SUBDIRECTORIO LLAMADO "EXCEL" 
		os.mkdir(os.path.join(path, "EXCEL"))  
		created = True
	else:
		raise Exception('Ya existen boletines de esta sección en el directorio seleccionado. Por favor, cambie el directorio o elimine los archivos conflictivos.')
	return created

def create_pdfs_boletin(path: str, loadingScreen):
	"""Crea los archivos PDF de los boletines de los estudiantes.

	Los archivos PDF son ejecutados por una función en un hilo secundario para
	limitar problemas de rendimiento. 

	Attributes:
	path (str): Ruta del directorio principal.
	loadingScreen (LoadingScreen) : Pantalla de carga del proceso de generar archivos.
	"""
	# ARCHIVOS DENTRO DEL SUBDIRECTORIO "EXCEL"
	files = [i for i in os.listdir(os.path.join(path, 'EXCEL')) if i.endswith(".xlsx")]

	# ASIGNA LA CANTIDAD DE ARCHIVOS A PROCESAR 
	loadingScreen.assign_file_amount(len(files))
	for file in files:
		# RUTA DE SALIDA DEL ARCHIVO PDF
		output_file = os.path.join(path, 'PDF', f"{file.split('.')[0]}.pdf") 
		# RUTA DE ENTRADA DEL ARCHIVO EXCEL
		input_file = os.path.join(path, 'EXCEL', file) 
		
		input_file = input_file.replace("/", '\\')
		output_file = output_file.replace('/', '\\')

		# EJECUTA EL PROCESO DE EXPORTACIÓN EN SEGUNDO PLANO
		run_export_in_background(input_file, output_file)
		loadingScreen.progress_on_file(output_file)

def create_excel_boletin(student, school_year, subjects, mention, sheet_choiced, guide_teacher, date, path):
	"""Crea el archivo de Excel del boletín de un estudiante en la ruta indicada.

	
	Attributes:
	student (Student) : Instancia de la clase Student.
	school_year (str) : Año escolar.
	subjects (Subject) : Lista de instancias de la clase Subject.
	mention (str) : Mención del estudiante.
	sheet_choiced (str) : Año y sección del estudiante.
	guide_teacher (str) : Profesor guía del estudiante.
	date (str) : Fecha de entrega del boletín.
	path (str) : Ruta del directorio principal."""
	row = 9
	def_general = 0
	boletin = load_workbook('./Resource/BOLETIN.xlsx', data_only=True)
	sheet = boletin.active
	if student.name!= str:
		# AÑO ESCOLAR
		sheet['C4'].value += school_year
		# NOMBRE COMPLETO
		sheet['A5'].value += f"{student.name} {student.last_name}"
		# CEDULA DEL ESTUDIANTE
		sheet['I5'].value += student.cedula
		# MENCIÓN 
		sheet['A6'].value += mention
		# AÑO-SECCIÓN
		sheet['D6'].value += sheet_choiced
		# PROFESOR GUÍA
		sheet['F6'].value += guide_teacher
		# FECHA DE ENTREGA
		sheet['J6'].value += date

		for subject in subjects:
			if row >19:
				break
			else:
				# MATERIAS
				sheet['A'+ str(row)].value = subject.name.upper()

				student_data = student.subjects_performance[subject.name]
				# NOTA DEL MOMENTO I
				sheet['H'+ str(row)].value = str(student_data.moment_grades[0])
				# NOTA DEL MOMENTO II 
				sheet['I'+ str(row)].value = str(student_data.moment_grades[1])
				# NOTA DEL MOMENTO III 
				sheet['J'+ str(row)].value = str(student_data.moment_grades[2])
				# NOTA DEFINITIVA	
				sheet['K'+ str(row)].value = str(student_data.moment_grades[3]) 
				if student.subjects_performance[subject.name].moment_grades[3]!="**":
					# SUMATORIA DE LAS NOTAS DEFINITIVAS
					def_general+=student.subjects_performance[subject.name].moment_grades[3] 
			row+=1
		# PROMEDIO GENERAL DEL ESTUDIANTE
		sheet['K22'].value = str(round(def_general/len(subjects), 2))
			
		boletin.save(os.path.join(path, 'EXCEL', f'{student.cedula}.xlsx'))
		boletin.close()
	
