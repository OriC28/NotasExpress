# -*- coding: utf-8 -*-
"""
Archivo: WriteFile.py
Descripción: Este archivo contiene las funciones necesarias para crear los archivos de los boletines de los estudiantes.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""
# IMPORTANDO LIBRERÍAS NECESARIAS
from openpyxl import load_workbook
import os

# IMPORTANDO MÓDULOS LOCALES
from Lib.GetPDF import run_export_in_background

"""
Crea los directorios necesarios para almacenar los archivos de los boletines

@param path: Ruta del directorio principal.
@return created: True si se crearon los directorios, False en caso contrario.
"""
def create_folders(path: str):
	created = False
	if not os.path.exists(path):
		# CREA EL DIRECTORIO PRINCIPAL 
		os.mkdir(path)
		# CREA UN SUBDIRECTORIO LLAMADO "PDF" 
		os.mkdir(os.path.join(path, "PDF"))
		# CREA UN SUBDIRECTORIO LLAMADO "EXCEL" 
		os.mkdir(os.path.join(path, "EXCEL"))  
		created = True
	return created

"""
Crea los archivos PDF de los boletines de los estudiantes.

@param path: Ruta del directorio principal.
@param loadingScreen: Instancia de la clase LoadingScreen.
@param enableButtonFunc: Función que habilita el botón de "Generar Boletines".
"""
def create_pdfs_boletin(path: str, loadingScreen, enableButtonFunc=None):
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

		# ACTUALIZA LA ETIQUETA QUE INDICA EL ARCHIVO QUE SE ESTÁ PROCESANDO
		loadingScreen.get_file_direction_txt()
		# EJECUTA EL PROCESO DE EXPORTACIÓN EN SEGUNDO PLANO
		run_export_in_background(input_file, output_file, loadingScreen.progress_on_file)
"""
Crea el archivo de Excel del boletín de un estudiante.

@param student: Instancia de la clase Student.
@param school_year: Año escolar.
@param subjects: Lista de instancias de la clase Subject.
@param mention: Mención del estudiante.
@param sheet_choiced: Año y sección del estudiante.
@param guide_teacher: Profesor guía del estudiante.
@param date: Fecha de entrega del boletín.
@param path: Ruta del directorio principal.
"""
def create_excel_boletin(student, school_year, subjects, mention, sheet_choiced, guide_teacher, date, path):
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
			if row >17:
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
		sheet['K20'].value = str(round(def_general/len(subjects), 2))
			
		boletin.save(os.path.join(path, 'EXCEL', f'{student.cedula}.xlsx'))
		boletin.close()
	
