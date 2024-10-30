class Student:
    def __init__(self,cedula=str,name=str,last_name=str):
        self.cedula=cedula
        self.name=name
        self.last_name = last_name
    def __str__(self):
        return str(self.name)

    __repr__ = __str__

class Subject:
    def __init__(self, name=str):
        #Quiero colocar aquí un promedio general de cada matería, 
        # asi en los boletines se puede colocar el promedio general por asignatura
        self.name = name

class Gradings(Subject):
    def __init__(self, name=str, moment_grades=list):
        super().__init__(name)
        self.moment_grades = moment_grades



