"""
Archivo: main.py
Descripción: Módulo principal para la generación de boletines de calificaciones.
Autor: Oriana Colina
Fecha: 25 de febrero de 2025
"""

from PyQt6.QtWidgets import QApplication
from controllers.BoletinController import BoletinController

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

if __name__ == '__main__':
	app = QApplication([])
	controller = BoletinController()
	app.exec()

	