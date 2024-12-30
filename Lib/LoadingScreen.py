from PyQt6.QtWidgets import QWidget, QMessageBox
from PyQt6 import uic
from PyQt6.QtCore import Qt
from PyQt6 import QtCore
import re
	
class LoadingScreen(QWidget):
	finished_signal = QtCore.pyqtSignal()
	def __init__(self, messageBox, func):
		super().__init__()
		uic.loadUi('GUI/Loading.ui', self)
		self.messageBox = messageBox
		self.finished_signal.connect(func)
		self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint)
		#esta variable define la cantidad de archivos, para asignar su valor, se debe usar la funcion correspondiente
		self.fileAmount = 0
		self.file_counter_pos = self.file_counter.text().find('(0/0)')
		self.progressBar.valueChanged.connect(self.finished_loading)

	def get_file_direction_txt(self):
		#Texto formateado dentro de este label:
		#'<html><head/><body><p><span style=" font-size:10pt;">Archivo: //dirección</span></p></body></html>'
		current_text = self.file_direction.text().split('<html><head/><body><p><span style=" font-size:10pt;">')
		current_text = current_text[1].split('</span></p></body></html>')
		return current_text[0]
	
	def change_file_direction_txt(self, new_direction):
		current_text = self.get_file_direction_txt()
		new_text = self.file_direction.text().replace(current_text, 'Archivo: ' + new_direction)
		self.file_direction.setText(new_text)

	def get_file_counter_txt(self):
		#Texto formateado dentro de este label:
		#<html><head/><body><p align="center"><span style=" font-size:10pt;">(0/0)</span></p></body></html>
		current_text = self.file_counter.text().split('<html><head/><body><p align="center"><span style=" font-size:10pt;">')
		current_text = current_text[1].split('</span></p></body></html>')
		return current_text[0]

	def change_file_counter_txt(self, new_text):
		current_text = self.get_file_counter_txt()		
		self.file_counter.setText(self.file_counter.text().replace(current_text, new_text))

	def assign_file_amount(self, amount):
		self.fileAmount = amount
		self.change_file_counter_txt(f'(0/{amount})')

	def sum_to_progress_bar(self):
		self.progressBar.setValue(self.progressBar.value() + round(100/self.fileAmount))
		file_counter_current_txt = self.get_file_counter_txt()
		new_txt = file_counter_current_txt.replace(file_counter_current_txt[1], str(int(file_counter_current_txt[1]) + 1))
		self.change_file_counter_txt(new_txt)
	
	def progress_on_file(self, file_direction=''):
		#esta función realiza todas las funciones para señalar progreso en la barra (porcentaje, archivo realizado, cantidad de archivos listos)
		self.sum_to_progress_bar()
		self.change_file_direction_txt(file_direction)
	
	def finished_loading(self):
		if (self.progressBar.value() == round(100/self.fileAmount) * self.fileAmount):	
			self.close()
			self.messageBox.exec()
			self.finished_signal.emit()
			
			
