# -*- coding: utf-8 -*-
"""
Archivo: WriteFile.py
Descripción: Este módulo contiene las clases que se utilizan para la creación de los estudiantes y las materias.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""

"""
Clase que permite crear un estudiante con sus respectivas notas.

@param cedula (str): Cédula del estudiante.
@param name (str): Nombre del estudiante.
@param last_name (str): Apellido del estudiante.
"""
class Student:
    def __init__(self,cedula=str(''),name=str(''),last_name=str('')):
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
   
"""
Clase que permite crear las notas de los estudiantes en cada momento de evaluación.

@param name (str): Nombre de la materia.
@param moment_grades (list): Lista de notas de los estudiantes.
"""
class Gradings(Subject):
    def __init__(self, name: str, moment_grades=list):
        super().__init__(name)
        self.moment_grades = moment_grades
    def __str__(self):
        return str(self.moment_grades)

    __repr__ = __str__


