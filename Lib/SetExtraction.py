from Lib.Extraction import Extraction

def set_extraction(file_path, index_choiced):
	extraction = Extraction(file_path, index_choiced)
	first_table_position = extraction.find_start_end_table(14, 30)
	second_table_position = extraction.find_start_end_table(int(first_table_position[1])+1, 60)  

	students = extraction.get_student_data(first_table_position[0], first_table_position[1])
	students = extraction.save_student_notes(14, first_table_position, students)

	total_students = extraction.save_student_notes(int(second_table_position[0])-2, second_table_position, students)
	school_year = extraction.get_school_year()
	subjects = extraction.get_subjects(14) + extraction.get_subjects(int(second_table_position[0])-2)
	name_folder = extraction.sheets[index_choiced].replace('"', "")

	return total_students, school_year, subjects, name_folder