from PyQt6.QtWidgets import QWidget, QApplication, QPushButton, QMessageBox, QFileDialog
from PyQt6 import uic, QtGui
import os

from Lib.Extraction import Extraction
import Lib.WriteFile as wf
from Lib.SetExtraction import set_extraction
from Lib.CheckExcelProcess import check_excel_running

class Generator:
	def __init__(self):
		self.file_path = None
		self.save_path = None
		self.message = QMessageBox 
		self.program = uic.loadUi("GUI/GUI.ui")

		# ESTABLECIENDO EL LOGO
		imagen = QtGui.QPixmap("./Resource/LOGO.png")
		self.program.logo.setScaledContents(True)
		self.program.logo.resize(imagen.width(), imagen.height())
		self.program.logo.setPixmap(imagen)
		self.init_gui()

	def init_gui(self):
		self.program.SelectButton.clicked.connect(self.select_excel_file)
		self.program.SaveButton.clicked.connect(self.select_path_to_save)
		self.program.GenerateButton.clicked.connect(self.generate_notes)
		self.program.show()

	def select_path_to_save(self):
		try:
			dir_path = QFileDialog.getExistingDirectory(parent=self.program, caption="Select directory",
															directory=os.path.expanduser('~'),		
															options=QFileDialog.Option.DontUseNativeDialog)
			if dir_path:
				self.program.SaveButton.setEnabled(True)
				self.save_path = dir_path
				self.program.pathFolder.setText(self.save_path)

		except Exception as e:
			self.message.warning(self.program, "Warning", f"Error. {e}", QMessageBox.StandardButton.Ok)
			
	def select_excel_file(self):
		try:
			self.program.CbYearSection.clear()
			self.program.CbYearSection.setEnabled(False)
			self.program.pathFile.setText('')
			file_path, filter = QFileDialog.getOpenFileName(self.program, 'Open file', '', 'Excel files (*.xlsx)')
			if file_path:
				if self.show_dialog():
					self.file_path = file_path
					self.program.pathFile.setText(file_path)

					extraction = Extraction(self.file_path)

					self.program.CbYearSection.addItems(extraction.sheets)
					self.program.CbYearSection.setEnabled(True)
					del extraction
					self.program.SaveButton.setEnabled(True)
			else:
				self.message.warning(self.program, "Archivo no encontrado", f"No se ha seleccionado ningún archivo.", QMessageBox.StandardButton.Ok)

		except Exception as e:
			self.message.warning(self.program, "Warning", f"{e}", QMessageBox.StandardButton.Ok)

	def show_dialog(self):
		answer = self.message.question(self.program, "Question", "¿Está seguro de generar estos boletines?",
					QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
		return True if answer == 16384 else False	
			
	def generate_notes(self):
		if not check_excel_running():
			try:
				index_choiced = self.program.CbYearSection.currentIndex()
				if index_choiced!=-1:
					total_students, school_year, subjects = set_extraction(self.file_path, index_choiced)
					sheet_choiced_name = self.program.CbYearSection.currentText()
					mention = self.program.CbMention.currentText()
					guide_teacher = self.program.GuideTeacherEntry.text()
					date = self.program.DateEntry.text()
					
					name_folder = sheet_choiced_name.replace('"', "")
					if self.save_path:
						path = os.path.join(self.save_path, name_folder.replace(' ', "_"))
						if wf.create_folders(path):
							for student in total_students:
								wf.create_excel_boletin(student, school_year, subjects, mention, sheet_choiced_name, guide_teacher, date, path)
							wf.create_pdfs_boletin(path)
					else:
						self.message.warning(self.program, "Ruta destino no encontrada", "Error. No se ha seleccionado una ruta para guardar los boletines.", QMessageBox.StandardButton.Ok)
				else:
					self.message.warning(self.program, "Warning", "Por favor seleccione un archivo para generar los boletines.", QMessageBox.StandardButton.Ok)
			except Exception as e:
				self.message.warning(self.program, "Warning", f"{e}", QMessageBox.StandardButton.Ok)
		else:
			self.message.warning(self.program, "EXCEL.EXE está activo", "Se ha detectado que el programa Excel está abierto, por favor ciérrelo para continuar.", QMessageBox.StandardButton.Ok)

if __name__ == '__main__':
	app = QApplication([])
	generator = Generator()
	app.exec()