"""
Archivo: main.py
Descripción: Módulo principal para la generación de boletines de calificaciones.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

# IMPORTANDO LIBRERIAS NECESARIAS
from PyQt6.QtWidgets import QApplication, QMessageBox, QFileDialog, QTableWidget, QTableWidgetItem
from PyQt6 import uic, QtGui

import os
import re
import threading

# IMPORTANDO MÓDULOS LOCALES
from Lib.LoadingScreen import LoadingScreen
from Lib.Extraction import Extraction
import Lib.WriteFile as wf
from Lib.SetExtraction import set_extraction
from Lib.CheckExcelProcess import check_excel_running


class Generator:
	"""Clase para generar los boletines de calificaciones.

	Attributes:
	file_path (str): Ruta del archivo Excel.
	save_path (str): Ruta de guardado de los boletines.
	message (QMessageBox): Objeto de mensaje.
	table (QTableWidget): Tabla de estudiantes.
	program (QMainWindow): Programa principal."""

	def __init__(self):
		"""Inicialización de la clase Generator."""
		self.file_path = None
		self.save_path = None
		self.message = QMessageBox 
		self.table = QTableWidget
		self.program = uic.loadUi("GUI/GUI.ui")
		self.init_gui()
		
	def init_gui(self):
		"""Inicializa la interfaz general de usuario cargada anteriormente.
		
		Establece el logo de la interfaz, conecta una multitud de señales, y muestra la ventana."""
		# ESTABLECIENDO EL LOGO
		imagen = QtGui.QPixmap("./Resource/LOGO.png")
		self.program.logo.setScaledContents(True)
		self.program.logo.resize(imagen.width(), imagen.height())
		self.program.logo.setPixmap(imagen)
		self.program.setWindowIcon(QtGui.QIcon('Resource/icons/iconProgram.ico'))
		self.program.SelectButton.clicked.connect(self.select_excel_file)
		self.program.SaveButton.clicked.connect(self.select_path_to_save)
		self.program.GenerateButton.clicked.connect(self.generate_notes)
		self.program.CbYearSection.currentIndexChanged.connect(self.fill_table)
		self.program.show()

	def select_path_to_save(self):
		"""Obtiene la dirección deseada para guardar los archivos a generar ingresada por el usuario."""
		try:
			dir_path = QFileDialog.getExistingDirectory(parent=self.program, caption="Select directory",
															directory=os.path.expanduser('~'),		
															options=QFileDialog.Option.DontUseNativeDialog)
			if dir_path:
				self.program.SaveButton.setEnabled(True)
				self.save_path = dir_path
				self.program.pathFolder.setText(self.save_path)

		except Exception as e:
			self.message.warning(self.program, "Error", f"Ha ocurrido un error: {e}", QMessageBox.StandardButton.Ok)
			
	def select_excel_file(self):
		"""Obtiene la dirección del archivo Excel a ser procesado.
		
		En caso de que la dirección no exista, o sea inválido, arroja una excepción."""
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
		"""Rellena los campos de la tabla de la interfaz de usuario con los datos obtenidos
		de una de las hojas del archivo Excel insertado por el usuario.
		
		Esta función obtiene los datos de la sección elegida en la interfaz para rellenar la tabla."""
		try:
			i=0
			index_choiced = self.program.CbYearSection.currentIndex()
			if index_choiced < 0:
				index_choiced = 0

			data = set_extraction(self.file_path, index_choiced)
			
			if data != False:
				total_students = data[0]
				row_count = len(total_students) if len(total_students)>=8 else 8
				self.program.Table.clearContents()
				self.program.Table.setRowCount(row_count)

				self.program.CbMention.setEnabled(True)
				if self.program.CbMention.count() == 0:
					self.program.CbMention.addItems([
						'TRANSPORTE ACUÁTICO', 
						'METALMECÁNICA', 
						'METALMECÁNICA Y NAVAL', 
						'MECÁNICA DE MANTENIMIENTO INDUSTRIAL'
					])
				
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
			else:
				self.program.CbYearSection.clear()
				self.program.CbMention.clear()
				self.program.CbYearSection.setEnabled(False)
				self.program.CbMention.setEnabled(False)
				self.program.Table.clearContents()
				
		except Exception as e:
			self.message.warning(self.program, "Error", f"Ha ocurrido un error: {e}", QMessageBox.StandardButton.Ok)

	def show_dialog(self):
		"""Muestra un mensaje para asegurar que el usuario desea generar los boletines elegidos."""
		answer = self.message.question(self.program, "Question", "¿Está seguro de generar estos boletines?",
					QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
		return True if answer == 16384 else False	
	
	def validate_fields(self):
		"""Valida todos los campos de la interfaz de usuario.
		
		En caso de no cumplir con las validaciones, arroja una excepción en el punto donde se detecte la primera
		inconsistencia.
		
		Los campos validados son 'guide_teacher' (campo de Profesor guía - str) y 'date' (Fecha de expedición - str)
		
		Los campos son validados en base a las siguientes especificaciones:
		'guide_teacher' : Campo NO vacío, solo puede contener letras, acentos y espacios.
		'date' : Campo NO vacío, solo puede contener números y debe estar escrito en el formato dd/mm/aa"""

		guide_teacher = self.program.GuideTeacherEntry.text()
		date = self.program.DateEntry.text()

		if guide_teacher == '':
			raise Exception('El campo de Profesor guía debe ser llenado para la generación de los archivos.')
		
		if date == '': 
			raise Exception ('El campo de fecha de entrega debe ser llenado para la generación de los archivos.')
		
		if not re.match("^[a-zA-ZáéíóúÁÉÍÓÚüÜñÑ'\\s]+$", guide_teacher):
			raise Exception('Solo se pueden colocar letras como carácteres en el campo de profesor guía. Por favor, cambialos para continuar')
		
		elif not re.match("^([1-9]|[12][0-9]|3[01])\\/([1-9]|0[1-9]|1[0-2])\\/(19|20)\\d{2}$", date):
			raise Exception('Asegúrate de que la fecha está bien escrita, debe seguir el formato: "día/mes/año" .')

	
	def generate_notes(self):
		"""Genera las notas de la hoja del archivo Excel seleccionado.
		
		Para su funcionamiento es necesario un hilo secundario para la generación de los
		archivos PDF y la pantalla de carga.
		
		También, es necesario que la aplicación Excel esté cerrada para su funcionamiento
		(se arroja una advertencia en caso contrario), además de que los campos estén validados.
		"""
		self.loading = LoadingScreen(self.Finished_Generation_MSGBox(), self.Enable_GenerateButton)

		try:
			if check_excel_running():
				self.loading.close()
				raise Exception("Se ha detectado que el programa Excel está abierto, por favor ciérrelo para continuar.")

			index_choiced = self.program.CbYearSection.currentIndex()
			self.validate_fields()

			if index_choiced ==-1:
				self.loading.close()
				raise Exception("Si ha seleccionado un archivo incorrecto no será posible continuar. Por favor seleccione un archivo adecuado.")
			
			self.Enable_GenerateButton()
			self.loading.show()
			data = set_extraction(self.file_path, index_choiced)

			if data != False:
				self.Enable_GenerateButton()
				self.program.GenerateButton.setEnabled(False)
				total_students, school_year, subjects = data
				sheet_choiced_name = self.program.CbYearSection.currentText()
				mention = self.program.CbMention.currentText()
				guide_teacher = self.program.GuideTeacherEntry.text()
				date = self.program.DateEntry.text()
				name_folder = sheet_choiced_name.replace('"', "")
				
				if not self.save_path:
					self.loading.close()
					self.Enable_GenerateButton()
					raise Exception("Error. No se ha seleccionado una ruta para guardar los boletines.")
				
				path = os.path.join(self.save_path, name_folder.replace(' ', "_"))
				self.Enable_GenerateButton()

				if wf.create_folders(path):
					for student in total_students:
						#GENERACIÓN DE ARCHIVOS EXCEL
						wf.create_excel_boletin(student, school_year, subjects, mention, sheet_choiced_name, guide_teacher, date, path)
							
					#HILO SECUNDARIO PARA GENERAR ARCHIVOS PDF
					thread = threading.Thread(target=wf.create_pdfs_boletin,args=(path, self.loading))
					thread.start()
			else:
				self.program.CbYearSection.clear()
				self.program.CbMention.clear()
				self.program.CbYearSection.setEnabled(False)
				self.program.CbMention.setEnabled(False)
				self.program.Table.clearContents()
				self.loading.close()
				raise Exception("Error. El archivo seleccionado no es una sabana de notas válida.")
	
		except Exception as e:
			self.loading.close()
			self.message.warning(self.program, "Advertencia", f"Ha ocurrido un error: {e}", QMessageBox.StandardButton.Ok)
			

	def Finished_Generation_MSGBox(self):
		"""Genera un mensaje que indica que la generación de archivos
		ha finalizado correctamente."""
		msg_box = QMessageBox()
		msg_box.setIcon(QMessageBox.Icon.Information)
		msg_box.setText("Se han generado todos los archivos PDF correctamente.")
		msg_box.setWindowTitle("¡Generación completada!")
		msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
		return msg_box
	
	def Enable_GenerateButton(self):
		"""Habilita el botón de generación nuevamente."""
		self.program.GenerateButton.setEnabled(True)
	
if __name__ == '__main__':
	app = QApplication([])
	generator = Generator()
	app.exec()