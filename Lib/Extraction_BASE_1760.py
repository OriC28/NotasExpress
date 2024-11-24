from openpyxl import load_workbook
import re
from Lib.students import Student

COLUMNS = [["F","G","H","I"],["J","K","L","M"],["N","O","P","Q"],["R","S","T","U"],["V","W","X","Y"]]

class Extraction:
	def __init__(self, file_path, choiced=0):
		self.file_path = file_path
		self.choiced = choiced
		self.workbook = load_workbook(self.file_path, data_only=True)
		self.sheets = self.workbook.sheetnames
		self.sheet_choiced = self.workbook[self.sheets[choiced]]

	def find_start_end_table(self, aprox_start, aprox_end):
		for n in range(aprox_start, aprox_end):
			if re.findall(r'^V-|^CE-', str(self.sheet_choiced['C' + str(n)].value)):
				start = str(n)
				break
		for n in range(int(start), aprox_end):
			if self.sheet_choiced['C' + str(n)].value is None:
				end = str(n-1)
				break
		return [start,end]

	def get_student_data(self, start, end):
		students_list = []
		for row in range(start, end+1):
			new_student = Student()
			for column in 'CDE':
				row_data = self.sheet_choiced[str(column+str(row))].value
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

	def get_subjects(self, row):
		subjects = {}
		for i in self.sheet_choiced[row]:
			if i.value is not None and i.value!="Promedios":
				subjects.update({i.value: None})
		return subjects

	def get_notes(self, start, end, notes):
		moments = []
		note = []
		for i in self.sheet_choiced[start: end]:
			for j in i:
				if j.value is None:
					j.value = 0 
				note.append(round(j.value, 2))
				if len(note)==4:
					moments.append(note) 
					note = []
				if len(moments) == len(self.get_student_data(int(start[1:]), int(end[1:]))): 
					notes.append(moments)
					moments = []

	def save_notes_subjects(self, row_subjects, table_position):
		i = 0
		notes = []
		start = table_position[0]
		end = table_position[1]
		subjects = self.get_subjects(row_subjects)
		for block in COLUMNS[:len(subjects)]:
			self.get_notes(block[0]+start, block[-1]+end, notes)
			while(i<len(notes)): 
				if i == len(subjects): 
					break
				else:
					subjects[list(subjects.keys())[i]] = notes[i] 
					i+=1
		return subjects 

	def get_school_year(self):
		row_to_search = self.sheet_choiced["B11":"B11"][0][0].value
		year = re.search(r'(\d{4}-\d{4})', row_to_search)
		return year.group()

'''
e = Extraction("NOTAS.xlsx", 5) # LOS PARÁMETROS DE Extraction LOS PROPORCIONARÁ LA GUI

# POSICIONES APROXIMADAS PARA OBTENER LAS REALES
first_table_position = e.find_start_end_table(14, 30) 
second_table_position = e.find_start_end_table(31, 60)

# MATERIAS CON SUS RESPECTIVAS NOTAS
first_table_notes = e.save_notes_subjects(14, first_table_position)
second_table_notes = e.save_notes_subjects(35, second_table_position)

# DATOS DE LOS ESTUDIANTES (CEDULAS, NOMBRES Y APELLIDOS)
students_data = e.get_student_data(int(first_table_position[0]), int(first_table_position[1]))

'''