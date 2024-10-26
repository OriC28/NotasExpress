from PyQt6.QtWidgets import QWidget, QApplication, QPushButton, QMessageBox, QFileDialog
from PyQt6 import uic, QtGui
from threading import Thread
from Lib.Extraction import Extraction

class Generator:
	def __init__(self):
		self.file_path = None
		self.message = QMessageBox 
		self.program = uic.loadUi("GUI/GUI.ui")
		self.init_gui()

	def init_gui(self):
		self.program.AddButton.clicked.connect(self.select_excel_file)
		self.program.CbYearSection.currentIndexChanged.connect(self.get_notes)
		self.program.show()

	def select_excel_file(self):
		try:
			file_path, filter = QFileDialog.getOpenFileName(self.program, 'Open file', '', 'Excel files (*.xlsx)')
			if file_path:
				self.file_path = file_path
				self.program.DirectoryEntry.setText(file_path)

				extraction = Extraction(self.file_path)

				self.program.CbYearSection.addItems(extraction.sheets)
				self.program.CbYearSection.setEnabled(True)

		except Exception as e:
			print(e)

		finally:
			del extraction
			
	def get_notes(self):
		try:
			sheet_choiced = self.program.CbYearSection.currentIndex()
			extraction = Extraction(self.file_path, sheet_choiced)
	
			first_table_position = extraction.find_start_end_table(14, 30) 
			#second_table_position = self.e.find_start_end_table(31, 60)

			first_table_notes = extraction.save_notes_subjects(14, first_table_position)
			#second_table_notes = self.e.save_notes_subjects(35, second_table_position)

			print(first_table_notes)
		except Exception as e:
			print(e)

if __name__ == '__main__':
	app = QApplication([])
	generator = Generator()
	app.exec()