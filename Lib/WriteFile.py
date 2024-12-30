from openpyxl import load_workbook
import os

from Lib.GetPDF import run_export_in_background

def create_folders(path):
	created = False
	if not os.path.exists(path):
		os.mkdir(path) # CREA EL DIRECTORIO PRINCIPAL 
		os.mkdir(os.path.join(path, "PDF")) # CREA UN SUBDIRECTORIO LLAMADO "PDF"
		os.mkdir(os.path.join(path, "EXCEL"))  # CREA UN SUBDIRECTORIO LLAMADO "EXCEL"
		created = True
	return created

def create_pdfs_boletin(path, loadingScreen, enableButtonFunc=None):
	files = [i for i in os.listdir(os.path.join(path, 'EXCEL')) if i.endswith(".xlsx")] # ARCHIVOS DENTRO DEL SUBDIRECTORIO "EXCEL"
	loadingScreen.assign_file_amount(len(files))
	for file in files:
		output_file = os.path.join(path, 'PDF', f"{file.split('.')[0]}.pdf") # EJEMPLO: /DIRECTORIO_RAIZ/1ERO A/PDF/33725588.xlsx
		input_file = os.path.join(path, 'EXCEL', file) # EJEMPLO: /DIRECTORIO_RAIZ/1ERO A/EXCEL/32689581.xlsx
		
		input_file = input_file.replace("/", '\\')
		output_file = output_file.replace('/', '\\')
		loadingScreen.get_file_direction_txt()
		run_export_in_background(input_file, output_file, loadingScreen.progress_on_file)
	
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

				sheet['H'+ str(row)].value = str(student_data.moment_grades[0]) # NOTA DEL MOMENTO I
				sheet['I'+ str(row)].value = str(student_data.moment_grades[1]) # NOTA DEL MOMENTO II
				sheet['J'+ str(row)].value = str(student_data.moment_grades[2])	# NOTA DEL MOMENTO III
				sheet['K'+ str(row)].value = str(student_data.moment_grades[3]) # NOTA DEFINITIVA
				if student.subjects_performance[subject.name].moment_grades[3]!="**":
					def_general+=student.subjects_performance[subject.name].moment_grades[3] # SUMATORIA DE LAS NOTAS DEFINITIVAS
			row+=1
		# PROMEDIO GENERAL DEL ESTUDIANTE
		sheet['K20'].value = str(round(def_general/len(subjects), 2))
			
		boletin.save(os.path.join(path, 'EXCEL', f'{student.cedula}.xlsx'))
		boletin.close()
	
