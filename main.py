from PyQt6.QtWidgets import QWidget, QApplication, QPushButton, QMessageBox, QFileDialog
from PyQt6 import uic, QtGui

from Lib.Extraction import Extraction
import Lib.WriteFile
import Lib.SetExtraction

class Generator:
	def __init__(self):
		self.file_path = None
		self.message = QMessageBox 
		self.program = uic.loadUi("GUI/GUI.ui")

		# ESTABLECIENDO EL LOGO
		imagen = QtGui.QPixmap("./Resource/LOGO.png")
		self.program.logo.setScaledContents(True)
		self.program.logo.resize(imagen.width(), imagen.height())
		self.program.logo.setPixmap(imagen)

		self.init_gui()

	def init_gui(self):
		self.program.AddButton.clicked.connect(self.select_excel_file)
		self.program.GenerateButton.clicked.connect(self.generate_notes)
		self.program.show()
	
	def select_excel_file(self):
		try:
			self.program.CbYearSection.clear()
			self.program.CbYearSection.setEnabled(False)
			self.program.DirectoryEntry.setText('')
			file_path, filter = QFileDialog.getOpenFileName(self.program, 'Open file', '', 'Excel files (*.xlsx)')
			if file_path:
				if self.show_dialog():
					self.file_path = file_path
					self.program.DirectoryEntry.setText(file_path)

					extraction = Extraction(self.file_path)

					self.program.CbYearSection.addItems(extraction.sheets)
					self.program.CbYearSection.setEnabled(True)
					del extraction
			else:
				self.message.warning(self.program, "Archivo no encontrado", f"No se ha seleccionado ningún archivo.", QMessageBox.StandardButton.Ok)

		except Exception as e:
			self.message.warning(self.program, "Warning", f"{e}", QMessageBox.StandardButton.Ok)

	#def validate_fields(mention)

	def show_dialog(self):
		answer = self.message.question(self.program, "Question", "¿Está seguro de generar estos boletines?",
					QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
		return True if answer == 16384 else False	
			
	def generate_notes(self):
		try:
			index_choiced = self.program.CbYearSection.currentIndex()

			if index_choiced!=-1:
				
				# DATOS EXTRAÍDOS
				total_students, school_year, subjects, name_folder = Lib.SetExtraction.set_extraction(self.file_path, index_choiced)
				
				mention = self.program.CbMention.currentText()
				guide_teacher = self.program.GuideTeacherEntry.text()
				date = self.program.DateEntry.text()
				sheet_choiced_name = self.program.CbYearSection.currentText()

				Lib.WriteFile.create_folders(name_folder)

				for student in total_students:
					Lib.WriteFile.create_excel_boletin(student, school_year, subjects, mention, sheet_choiced_name, guide_teacher, date, name_folder)

				Lib.WriteFile.create_pdfs_boletin(name_folder)

			else:
				self.message.warning(self.program, "Warning", f"No ha seleccionado ningún archivo.", QMessageBox.StandardButton.Ok)

		except Exception as e:
			self.message.warning(self.program, "Warning", f"{e}", QMessageBox.StandardButton.Ok)

if __name__ == '__main__':
	app = QApplication([])
	generator = Generator()
	app.exec()