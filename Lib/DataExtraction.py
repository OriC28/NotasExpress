import re
from Lib.students import Student

COLUMNS = [["F","G","H","I"],["J","K","L","M"],["N","O","P","Q"],["R","S","T","U"],["V","W","X","Y"]]

def find_start_end_table(aprox_start, aprox_end, sheet_choiced):
	# OBTENIENDO EL INICIO DE LA TABLA SIN ENCABEZADO
	for n in range(aprox_start, aprox_end):
		if re.findall(r'^V-|^CE-', str (sheet_choiced['C' + str(n)].value)): 
			start = str(n)
			break

	# OBTENIENDO EL FINAL DE LA TABLA SIN ENCABEZADO
	for n in range(int(start), aprox_end):
		if sheet_choiced['C' + str(n)].value is None:
			end = str(n-1)
			break
	return [start,end]

# OBTENIENDO TODOS LOS DATOS DE LOS ESTUDIANTES (CEDULA, APELLIDOS Y NOMBRES)
def get_student_data(start, end, sheet_choiced):
	students_list = []
	for row in range(start, end+1):
		new_student = Student()
		for column in 'CDE':
			row_data = sheet_choiced[str(column+str(row))].value
			if row_data != '//' and row_data:
				match column:
					case 'C':
						new_student.cedula = row_data
					case 'D':
						new_student.last_name = row_data
					case 'E':
						new_student.name = row_data
		students_list.append(new_student)
	return students_list

# OBTENIENDO LOS NOMBRES DE LAS MATERIAS
def get_subjects(row, sheet_choiced):
	subjects = {}
	for i in sheet_choiced[row]:
		if i.value is not None and i.value!="Promedios":
			subjects.update({i.value: None})
	return subjects

# AGREGANDO LOS MOMENTOS Al ARRAY DE NOTAS
def get_notes(start, end, notes, sheet_choiced, students: Student):
	moments = []
	note = []
	for i in sheet_choiced[start: end]:
		for j in i:
			if j.value is None:
				j.value = 0 # SI ALGUNA CASILLA DE LAS NOTAS ESTÁ VACÍA
			note.append(round(j.value, 2))
			if len(note)==4:
				moments.append(note) # AGREGAR CUANDO ESTÉN LAS NOTAS DE LOS MOMENTOS Y LA DEFINITIVA
				note = []
			if len(moments) == len(students): # EL TOTAL DE LAS NOTAS POR MOMENTOS EQUIVALE AL NÚMERO DE ESTUDIANTES
				notes.append(moments)
				moments = []

# GUARDANDO LAS NOTAS EN SU RESPECTIVA MATERIA
def save_notes(subjects, table_position, sheet_choiced, students: Student):
	i = 0
	notes = []
	start = table_position[0]
	end = table_position[1]
	for block in COLUMNS[:len(subjects)]:
		get_notes(block[0]+start, block[-1]+end, notes, sheet_choiced, students)
		while(i<len(notes)): 
			if i == len(subjects): 
				break
			else:
				subjects[list(subjects.keys())[i]] = notes[i] #AGREGANDO LAS NOTAS A CADA MATERIA
				i+=1 # IR A LA SIGUIENTE MATERIA

def get_school_year(sheet_choiced):
	row_to_search = sheet_choiced["B11":"B11"][0][0].value
	year = re.search(r'(\d{4}-\d{4})', row_to_search)
	return year.group()