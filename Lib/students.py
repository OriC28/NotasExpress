# -*- coding: utf-8 -*-
"""
Archivo: WriteFile.py
Descripción: Este módulo contiene las clases que se utilizan para la creación de los estudiantes y las materias.
Autor: Oriana Colina, Carlos Noguera, Genesys Alvarado, Ángel Colina y María Quevedo.
Fecha: 15 de enero de 2025
"""

# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class Student:
    """Clase de tipo estudiante, que almacena los datos correspondientes
    con sus respectivas notas.
    
    
    Attributes:
        cedula (str): [Cédula]
        name (str) : [Nombre]
        last_name (str): [Apellido]
        subjects_performance (dict): [Notas]
        """
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
    """Clase que almacena el nombre de una asignatura.

    Attributes:
    name (str): Nombre de la materia.
    """
    def __init__(self, name):
        self.name = name
   

class Gradings(Subject):
    """Clase que permite crear las notas de los estudiantes en cada momento de evaluación.


    Attributes:
    name (str): Nombre de la materia.
    moment_grades (list): Lista de notas de los estudiantes.
    """
    def __init__(self, name: str, moment_grades=list):
        super().__init__(name)
        self.moment_grades = moment_grades
    def __str__(self):
        return str(self.moment_grades)

    __repr__ = __str__


