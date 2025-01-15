# -*- coding: utf-8 -*-
"""
Archivo: CheckExcelProcess.py
Descripción: Módulo que permite verificar si el proceso de Excel se encuentra en ejecución.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""
# IMPORTANDO LIBRERIAS NECESARIAS
import psutil

"""
Permite verificar si el proceso de Excel se encuentra en ejecución.
    
returns (bool): True si el proceso de Excel se encuentra en ejecución, False en caso contrario.
"""
def check_excel_running():
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
        