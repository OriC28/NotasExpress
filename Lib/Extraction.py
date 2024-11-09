from openpyxl import load_workbook
import re
from Lib.students import Student, Gradings, Subject

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
		for row in range(int(start), int(end)+1):
			new_student = Student()
			for column in 'CDE':
				cell_data = self.sheet_choiced[str(column+str(row))].value
				if cell_data!='**' and cell_data:
					match column:
						case'C':
							new_student.cedula = cell_data
						case 'D':
							new_student.last_name = cell_data
						case'E':
							new_student.name = cell_data
			students_list.append(new_student)
		return students_list

	def get_subjects(self, row):
		subjects = []
		for i in self.sheet_choiced[row]:
			if i.value is not None and i.value!="Promedios":
				subjects.append(Subject(i.value))
		return subjects

	def get_notes(self, row, block):
		notes = []
		for letter in block:
			if self.sheet_choiced[letter + str(row)].value is None:
				note = self.sheet_choiced[letter + str(row)].value  = '**'
			else:
				note = int(self.sheet_choiced[letter + str(row)].value)
			notes.append(note)
		return notes


	def save_student_notes(self, row_subjects, table_position, students):
		i = 0
		notes = []
		start = table_position[0]
		end = table_position[1]
		subjects = self.get_subjects(row_subjects)
		for i in range (int(start), int(end)+1):
			j = 0
			for subject in subjects:
				students[i-int(start)].subjects_performance[subject.name] = Gradings(subject.name, self.get_notes(i, COLUMNS[j]))
				j += 1
		return students 

	def get_school_year(self):
		row_to_search = self.sheet_choiced["B11":"B11"][0][0].value
		year = re.search(r'(\d{4}-\d{4})', row_to_search)
		return year.group()

