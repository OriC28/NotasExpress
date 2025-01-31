# -*- coding: utf-8 -*-
"""
Archivo: GetPDF.py
Descripción: Módulo que permite exportar un archivo de Excel a PDF en un hilo separado.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

# IMPORTANDO LIBRERIAS NECESARIAS
from win32com import client
from PyQt6 import QtCore
import threading
import pythoncom

"""
Clase para enviar señales entre hilos.

@attr signal (object): Señal para enviar mensajes entre hilos.
"""
class thread_signal(QtCore.QObject):
	"""Clase para enviar señales entre hilos, se basa en las señales de la clase QObject.


	Attributes:
	signal (pyqtSignal): Señal para enviar mensajes entre hilos.
	"""
	signal = QtCore.pyqtSignal(str)


class pdf_thread(threading.Thread):
	"""Clase para exportar un archivo de Excel a PDF en un hilo separado.


	Attributes:
	thread_finished (thread_signal): Señal para indicar que el hilo ha finalizado.
	input_path (str): Ruta del archivo Excel.
	output_path (str): Ruta del archivo PDF.
	"""
	def __init__(self, input_path='', output_path=''):
		"""Inicialización de la clase pdf_thread.
		

		Attributes:
		input_path (str) : Ruta del archivo Excel.
		output_path (str) : Ruta del archivo PDF."""
		threading.Thread.__init__(self)
		self.thread_finished = thread_signal()
		self.input_path = input_path
		self.output_path = output_path
	
	def run(self):
		"""Método para exportar un archivo de Excel a PDF en un hilo separado.
		
		Es necesario que no esté abierta la aplicación Excel para su funcionamiento correcto"""
		pythoncom.CoInitialize()
		try:
			# ABRIR LA APLICACION EXCEL
			excel = client.Dispatch("Excel.Application")
			# ABRIR EL ARCHIVO EXCEL 
			sheets = excel.Workbooks.Open(self.input_path)  
			sheet = sheets.Worksheets[0]
			excel.DisplayAlerts = False

			# EXPORTAR EL ARCHIVO EXCEL A PDF
			sheet.ExportAsFixedFormat(0, self.output_path)

		except Exception as e:
			print(e)

		finally:
			# CERRAR LA APLICACIÓN EXCEL
			excel.Quit()
			pythoncom.CoUninitialize()

def run_export_in_background(input_path: str, output_path: str):
	"""Permite exportar un archivo de Excel a PDF en un hilo separado.

	Solo es posible ejecutar un hilo a la vez.
	Attributes:
	input_path (str): Ruta del archivo Excel.
	output_path (str): Ruta del archivo PDF.
	"""
	thread = pdf_thread(input_path, output_path)
	thread.start()
	thread.join()

