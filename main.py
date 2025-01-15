# -*- coding: utf-8 -*-
"""
Archivo: main.py
Descripción: Módulo principal para la generación de boletines de calificaciones.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""
# IMPORTANDO LIBRERIAS NECESARIAS
from PyQt6.QtWidgets import QApplication, QMessageBox, QFileDialog, QTableWidget, QTableWidgetItem, QMenu, QSplashScreen
from PyQt6 import uic, QtGui
import os
import re

# IMPORTANDO MÓDULOS LOCALES
from Lib.LoadingScreen import LoadingScreen
from Lib.Extraction import Extraction
import Lib.WriteFile as wf
from Lib.SetExtraction import set_extraction
from Lib.CheckExcelProcess import check_excel_running

"""
Clase para generar los boletines de calificaciones.

@attr file_path (str): Ruta del archivo Excel.
@attr save_path (str): Ruta de guardado de los boletines.
@attr message (object): Objeto de mensaje.
@attr table (object): Tabla de estudiantes.
@attr program (object): Programa principal.
"""
class Generator:
	def __init__(self):
		self.file_path = None
		self.save_path = None
		self.message = QMessageBox 
		self.table = QTableWidget
		self.program = uic.loadUi("GUI/GUI.ui")

		# ESTABLECIENDO EL LOGO
		imagen = QtGui.QPixmap("./Resource/LOGO.png")
		self.program.logo.setScaledContents(True)
		self.program.logo.resize(imagen.width(), imagen.height())
		self.program.logo.setPixmap(imagen)
		self.program.setWindowIcon(QtGui.QIcon('Resource/icons/iconProgram.ico'))
		self.init_gui()
		
	def init_gui(self):
		self.program.SelectButton.clicked.connect(self.select_excel_file)
		self.program.SaveButton.clicked.connect(self.select_path_to_save)
		self.program.GenerateButton.clicked.connect(self.generate_notes)
		#self.setup_table_context()
		self.program.CbYearSection.currentIndexChanged.connect(self.fill_table)
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
	
	def fill_table (self):
		try:
			index_choiced = self.program.CbYearSection.currentIndex()
			if index_choiced < 0:
				index_choiced = 0
			total_students, school_year, subjects = set_extraction(self.file_path, index_choiced)
			row_count = len(total_students) if len(total_students)>=8 else 8
			self.program.Table.clearContents()
			self.program.Table.setRowCount(row_count)
			i=0

			for student in total_students:
				if (student.name):
					sum_notes = 0
					
					for subject in student.subjects_performance.keys():
						notes = student.subjects_performance[subject]
						sum_notes += notes.moment_grades[3]

					average_score = round (sum_notes/len(student.subjects_performance.keys()),2)
					full_item =[str(i+1),str(student.cedula), str(student.name), str(student.last_name),str(average_score)]
					j = 0
					for data in full_item:
						self.program.Table.setItem(i,j,QTableWidgetItem(data))
						j +=1
					i += 1
		except Exception as e:
			self.message.warning(self.program, "Warning", f"{e}", QMessageBox.StandardButton.Ok)
	
	def setup_table_context(self):
		action = self.program.Table.addAction('Generar solo un PDF')
		action.triggered.connect(self.generate_individual_pdf)

	def generate_individual_pdf (self, event):
		#get selected row
		row = self.program.Table.rowAt(event.pos().y())
		print (row)

	def show_dialog(self):
		answer = self.message.question(self.program, "Question", "¿Está seguro de generar estos boletines?",
					QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
		return True if answer == 16384 else False	
	
	def validate_fields(self):
		guide_teacher = self.program.GuideTeacherEntry.text()
		if guide_teacher == '': raise Exception('El campo de Profesor guía debe ser llenado para la generación de los archivos.')
		date = self.program.DateEntry.text()
		if date == '': raise Exception ('El campo de fecha de entrega debe ser llenado para la generación de los archivos.')
		if not re.match('^[a-zA-Z\\s]+$', guide_teacher):
			raise Exception('Solo se pueden colocar letras como carácteres en el campo de profesor guía. Por favor, cambialos para continuar')
		elif not re.match("^([1-9]|[12][0-9]|3[01])\\/([1-9]|0[1-9]|1[0-2])\\/(19|20)\\d{2}$", date):
			raise Exception('Asegúrate de que la fecha está bien escrita, debe seguir el formato: "día/mes/año" .')
			
		
	def generate_notes(self):
		if not check_excel_running():
			self.loading = LoadingScreen(self.Finished_Generation_MSGBox(), self.Enable_GenerateButton)
			self.program.GenerateButton.setEnabled(False)
			try:
				index_choiced = self.program.CbYearSection.currentIndex()
				self.validate_fields()
				if index_choiced!=-1:
					self.loading.show()
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
							wf.create_pdfs_boletin(path, self.loading, self.Enable_GenerateButton)
					else:
						self.loading.close()
						self.message.warning(self.program, "Ruta destino no encontrada", "Error. No se ha seleccionado una ruta para guardar los boletines.", QMessageBox.StandardButton.Ok)
						self.Enable_GenerateButton()
				else:
					self.message.warning(self.program, "Warning", "Por favor seleccione un archivo para generar los boletines.", QMessageBox.StandardButton.Ok)
					self.Enable_GenerateButton()
			except Exception as e:
				self.loading.close()
				self.message.warning(self.program, "Warning", f"{e}", QMessageBox.StandardButton.Ok)
				self.Enable_GenerateButton()
		else:
			self.message.warning(self.program, "EXCEL.EXE está activo", "Se ha detectado que el programa Excel está abierto, por favor ciérrelo para continuar.", QMessageBox.StandardButton.Ok)
			self.loading.close()
			self.Enable_GenerateButton()
	
	def Finished_Generation_MSGBox(self):
		msg_box = QMessageBox()
		msg_box.setIcon(QMessageBox.Icon.Information)
		msg_box.setText("Se han generado todos los archivos PDF correctamente.")
		msg_box.setWindowTitle("¡Generación completada!")
		msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
		return msg_box
	
	def Enable_GenerateButton(self):
		self.program.GenerateButton.setEnabled(True)
	
if __name__ == '__main__':
	app = QApplication([])
	generator = Generator()
	app.exec()