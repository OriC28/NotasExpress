from win32com import client
import threading
import pythoncom
from PyQt6 import QtCore
class thread_signal (QtCore.QObject):
	signal = QtCore.pyqtSignal(str)

class pdf_thread (threading.Thread):
	def __init__(self, input_path='', output_path=''):
		threading.Thread.__init__(self)
		self.thread_finished = thread_signal()
		self.input_path = input_path
		self.output_path = output_path

	def run(self):
		pythoncom.CoInitialize()
		try:
			excel = client.Dispatch("Excel.Application") # ABRIR LA APLICACION EXCEL
			sheets = excel.Workbooks.Open(self.input_path)  # ABRIR WOORKBOOK
			sheet = sheets.Worksheets[0]
			excel.DisplayAlerts = False

			sheet.ExportAsFixedFormat(0, self.output_path)
			self.thread_finished.signal.emit(self.output_path)
		except Exception as e:
			print(e)

		finally:
			#sheets.Close(SaveChanges=False)
			excel.Quit()
			pythoncom.CoUninitialize()

def run_export_in_background(input_path, output_path, func):
	thread = pdf_thread(input_path, output_path)
	thread.thread_finished.signal.connect(func)
	thread.start()


