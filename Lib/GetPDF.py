from win32com import client
import threading
import pythoncom

def save_pdf(input_path, output_path):
	pythoncom.CoInitialize()
	try:
		excel = client.Dispatch("Excel.Application") # ABRIR LA APLICACION EXCEL
		sheets = excel.Workbooks.Open(input_path)  # ABRIR WOORKBOOK
		number_sheets = sheets.Worksheets.Count
		sheet = sheets.Worksheets[0]
		excel.DisplayAlerts = False

		sheet.ExportAsFixedFormat(0, output_path)

	except Exception as e:
		print(e)

	finally:
		#sheets.Close(SaveChanges=False)
		excel.Quit()
		pythoncom.CoUninitialize()

def run_export_in_background(input_path, output_path):
    thread = threading.Thread(target=save_pdf, args=(input_path, output_path))
    thread.start()

