import psutil

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
        