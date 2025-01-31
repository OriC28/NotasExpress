# -*- coding: utf-8 -*-
"""
Archivo: LoadingScreen.py
Descripción: Módulo que permite incorporar una ventana de carga durante la generación de boletines.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

from PyQt6.QtWidgets import QWidget
from PyQt6 import uic
from PyQt6.QtCore import Qt
from PyQt6 import QtCore
import re

class LoadingScreen(QWidget):
	"""Clase que define una pantalla de carga para la generación de los archivos.
	
	Está compuesta de una barra de progreso, una cantidad total de archivos
	(incluyendo cuantos han sido generados) y la dirección del último archivo generado.

	Hereda de la clase QWidget para su funcionamiento.

	Se recomienda ejecutar o manejar los métodos en un hilo secundario para evitar problemas de optimización.
	
	Attributes:
	messageBox (QMessageBox) : Mensaje final a ejecutar cuando se finalice la carga.
	finished_signal (pyqtSignal) : Señal emitida cuando finalice la carga. Ejecuta una función una vez esto sucede.
	func (function) : Función a ser ejecutada una vez se emita la señal de finalización.
	file_amount (int) : Cantidad de archivos a ser generados, este valor es usado para mostrar progreso en la pantalla."""
	
	finished_signal = QtCore.pyqtSignal()
	
	def __init__(self, messageBox, func):
		"""Inicialización de la clase LoadingScreen
		

		Attributes:
		messageBox (QMessageBox) : Mensaje final a ejecutar cuando se finalice la carga.
		func (function) : Función a ser ejecutada una vez se emita la señal de finalización.
		"""
		super().__init__()
		uic.loadUi('GUI/Loading.ui', self)
		self.messageBox = messageBox
		self.finished_signal.connect(func)
		self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint)
		self.fileAmount = 0
		self.progressBar.valueChanged.connect(self.finished_loading)

	def get_file_direction_txt(self):
		"""Obtiene el texto actual mostrado en el QLabel 'file_direction', 
		el cual, indica la dirección del archivo generado recientemente.
		
		
		Returns:
		str : Texto actual del QLabel 'file_direction'."""
		#Texto formateado dentro de este label:
		#'<html><head/><body><p><span style=" font-size:10pt;">Archivo: //dirección</span></p></body></html>'
		current_text = self.file_direction.text().split('<html><head/><body><p><span style=" font-size:10pt;">')
		current_text = current_text[1].split('</span></p></body></html>')
		return current_text[0]
	
	def change_file_direction_txt(self, new_direction):
		"""Cambia el texto actual mostrado en el QLabel 'file_direction',
		el cual, indica la dirección del archivo generado recientemente.
		
		
		Attributes:
		new_direction (str) : Nueva dirección a mostrar."""

		current_text = self.get_file_direction_txt()
		new_text = self.file_direction.text().replace(current_text, 'Archivo: ' + new_direction)
		self.file_direction.setText(new_text)

	def get_file_counter_txt(self):
		"""Obtiene el texto actual mostrado en el QLabel 'file_counter', 
		el cual, indica la cantidad de archivos ya generados.
		
		
		Returns:
		str : Texto actual del QLabel 'file_counter'."""
		#Texto formateado dentro de este label:
		#<html><head/><body><p align="center"><span style=" font-size:10pt;">(0/0)</span></p></body></html>
		current_text = self.file_counter.text().split('<html><head/><body><p align="center"><span style=" font-size:10pt;">')
		current_text = current_text[1].split('</span></p></body></html>')
		return current_text[0]

	def change_file_counter_txt(self, new_text):
		"""Cambia el texto actual mostrado en el QLabel 'file_counter',
		el cual, indica la cantidad de archivos ya generados.
		
		
		Attributes:
		new_text (str) : Nueva cantidad a mostrar.
		"""
		current_text = self.get_file_counter_txt()		
		self.file_counter.setText(self.file_counter.text().replace(current_text, new_text))

	def assign_file_amount(self, amount):
		"""Función correspondiente para asignar el valor de la variable 'fileAmount'.
		

		Attributes:
		amount (int) : Cantidad nueva."""
		self.fileAmount = amount
		self.change_file_counter_txt(f'(0/{amount})')

	def sum_to_progress_bar(self):
		"""Aumenta el porcentaje de la barra de progreso.
		
		Utiliza un valor basado en la cantidad de archivos a generar.
		
		El valor es igual al redondeo de: (100/Cantidad de archivos)
		"""
		new_value = self.progressBar.value() + round(100/self.fileAmount)
		new_value = 100 if new_value > 100 else new_value
		self.progressBar.setValue(new_value)
		file_counter_current_txt = self.get_file_counter_txt()
		match = re.match(r'\((\d+)/(\d+)\)', file_counter_current_txt)
		current_number = int(match.group(1))
		total_number = int(match.group(2))
		new_txt = f'({current_number + 1}/{total_number})'
		self.change_file_counter_txt(new_txt)

	
	def progress_on_file(self, file_direction=''):
		"""Realiza todas las funciones para señalar progreso en la barra (porcentaje, archivo realizado, cantidad de archivos listos)."""
		self.sum_to_progress_bar()
		self.change_file_direction_txt(file_direction)
	
	def finished_loading(self):
		"""Esta función es ejecutada nativamente por cada vez que cambia el progreso de la pantalla.
		
		Si el progreso es completado, entonces finaliza la ventana."""
		if (self.progressBar.value() == 100 or
	  	self.get_file_counter_txt() == f'({self.fileAmount}/{self.fileAmount})'):	
			self.close_window()

	def close_window(self):
		"""Ejecuta un mensaje y señal de finalización al cerrar la pantalla."""
		self.close()
		self.messageBox.exec()
		self.finished_signal.emit()
			