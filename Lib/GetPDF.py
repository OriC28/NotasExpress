# -*- coding: utf-8 -*-
"""
Archivo: GetPDF.py
Descripción: Módulo que permite exportar un archivo de Excel a PDF en un hilo separado.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""
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
	signal = QtCore.pyqtSignal(str)

"""
Clase para exportar un archivo de Excel a PDF en un hilo separado.

@attr thread_finished (object): Señal para indicar que el hilo ha finalizado.
@attr input_path (str): Ruta del archivo Excel.
@attr output_path (str): Ruta del archivo PDF.
"""
class pdf_thread(threading.Thread):
	
	def __init__(self, input_path='', output_path=''):
		threading.Thread.__init__(self)
		self.thread_finished = thread_signal()
		self.input_path = input_path
		self.output_path = output_path

	"""
	Método para exportar un archivo de Excel a PDF en un hilo separado.
	"""
	def run(self):
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
			# ENVIAR SEÑAL DE FINALIZACIÓN	
			self.thread_finished.signal.emit(self.output_path)
		except Exception as e:
			print(e)

		finally:
			# CERRAR LA APLICACIÓN EXCEL
			excel.Quit()
			pythoncom.CoUninitialize()

"""
Permite exportar un archivo de Excel a PDF en un hilo separado.

@param input_path (str): Ruta del archivo Excel.
@param output_path (str): Ruta del archivo PDF.
@param func (function): Función a ejecutar al finalizar el hilo.
"""
def run_export_in_background(input_path: str, output_path: str, func: callable):
	thread = pdf_thread(input_path, output_path)
	thread.thread_finished.signal.connect(func)
	thread.start()
	thread.join()


