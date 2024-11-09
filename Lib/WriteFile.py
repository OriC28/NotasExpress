from openpyxl import load_workbook
import os

from Lib.Extraction import Extraction
from Lib.GetPDF import run_export_in_background

def create_folders(name_folder):
	directory = f"./{name_folder}"
	if not os.path.exists(directory):
		os.mkdir(directory) # CREA EL DIRECTORIO PRINCIPAL 
		os.mkdir(f"{directory}/PDF") # CREA UN SUBDIRECTORIO LLAMADO "PDF"
		os.mkdir(f"{directory}/EXCEL")  # CREA UN SUBDIRECTORIO LLAMADO "EXCEL"

def create_pdfs_boletin(name_folder):
	files = [i for i in os.listdir(name_folder + "/EXCEL") if i.endswith(".xlsx")] # ARCHIVOS DENTRO DEL SUBDIRECTORIO "EXCEL"
	for file in files:
		out_file = f"{os.getcwd()}/{name_folder}/PDF/{file.split('.')[0]}" # EJEMPLO: /DIRECTORIO_RAIZ/1ERO A/PDF/33725588.xlsx
		input_file = f"{os.getcwd()}/{name_folder}/EXCEL/{file}" # EJEMPLO: /DIRECTORIO_RAIZ/1ERO A/EXCEL/32689581.xlsx
		
		run_export_in_background(input_file, f'{out_file}.pdf')

def create_excel_boletin(student, school_year, subjects, mention, sheet_choiced, guide_teacher, date, name_folder):
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

				sheet['H'+ str(row)].value = str(student_data.moment_grades[0]) # NOTA DEL MOMENTO I
				sheet['I'+ str(row)].value = str(student_data.moment_grades[1]) # NOTA DEL MOMENTO II
				sheet['J'+ str(row)].value = str(student_data.moment_grades[2])	# NOTA DEL MOMENTO III
				sheet['K'+ str(row)].value = str(student_data.moment_grades[3]) # NOTA DEFINITIVA
				if student.subjects_performance[subject.name].moment_grades[3]!="**":
					def_general+=student.subjects_performance[subject.name].moment_grades[3] # SUMATORIA DE LAS NOTAS DEFINITIVAS
			row+=1
		# PROMEDIO GENERAL DEL ESTUDIANTE
		sheet['K20'].value = str(round(def_general/len(subjects), 2))
			
		boletin.save(f'{name_folder}/EXCEL/{student.cedula}.xlsx')
		boletin.close()
	
