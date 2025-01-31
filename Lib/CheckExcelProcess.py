# -*- coding: utf-8 -*-
"""
Archivo: CheckExcelProcess.py
Descripción: Módulo que permite verificar si el proceso de Excel se encuentra en ejecución.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""
# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

# IMPORTANDO LIBRERIAS NECESARIAS
import psutil


def check_excel_running():
    """Permite verificar si el proceso de Excel se encuentra en ejecución.
        
    Returns: 
    bool : True si el proceso de Excel se encuentra en ejecución, False en caso contrario.
    """
    running = False
    for proc in psutil.process_iter():
        try:
            pinfo = proc.as_dict(attrs=['name'])
            if pinfo['name'] == 'EXCEL.EXE':
                running = True
                break
        except psutil.NoSuchProcess:
            pass   
    return running     

