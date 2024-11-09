class Student:
    def __init__(self,cedula=str,name=str,last_name=str):
        self.cedula=cedula
        self.name=name
        self.last_name = last_name
        self.subjects_performance = {}
    def __str__(self):
        return str(self.name + "-" + str(self.subjects_performance))
    def set_value(self, key, value):
        self.my
    __repr__ = __str__

class Subject:
    def __init__(self, name):
        #Quiero colocar aquí un promedio general de cada matería, 
        # asi en los boletines se puede colocar el promedio general por asignatura
        self.name = name
   

class Gradings(Subject):
    def __init__(self, name, moment_grades=list):
        super().__init__(name)
        self.moment_grades = moment_grades
    def __str__(self):
        return str(str(self.moment_grades))

    __repr__ = __str__


