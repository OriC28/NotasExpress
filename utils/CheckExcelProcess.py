import psutil

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

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

