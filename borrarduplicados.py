# -*- coding: utf-8 -*-


import pandas as pd
import os
import xlsxwriter
import xlrd
import sys
import time
import six

"""El objetivo del programa es borrar los archivos Excel que tengan hojas duplicadas.
   El programa va a listar todos los archivos con extensión xlsx del directorio donde se encuentre el programa
   Se elige un documento y el programa elimina los registros que estén duplicados en cada hoja del archivo"""

DIRECTORIO_ACTUAL = os.path.dirname(os.path.realpath(__file__))

def recorrer_duplicados(directorio,n):
    lista_duplicados = []
    sin_duplicados = []
    i = 0
    j = 0
    for tipo_archivo in directorio:
        if ("xlsx" in tipo_archivo) and ("sin duplicar" not in tipo_archivo):
            print(tipo_archivo)
            lista_duplicados.append(tipo_archivo)
            i += 1
    if i > 0: 
        print("se encontaron:", i, "archivos para duplicar \n")
    print("-----------------------------------------------------\n")
    for tipo_archivo in directorio:
        if ("xlsx" in tipo_archivo) and ("sin duplicar" in tipo_archivo or "duplicar" in tipo_archivo or "duplicado" in tipo_archivo):
            print(tipo_archivo)
            sin_duplicados.append(tipo_archivo)
            j += 1

    if j > 0:
        print("Se encontraron", j, "archivos sin duplicar \n")
        
    if len(lista_duplicados + sin_duplicados) == 0: 
        print("NO HAY ARCHIVOS EN",DIRECTORIO_ACTUAL,"\n")
        print("Recordá añadir archivos al directorio\n")
        return n
    elif len(lista_duplicados) == 0: 
        print("No se encontraron archivos para duplicar.en: ", DIRECTORIO_ACTUAL)
    elif len(sin_duplicados) == 0: 
        print("No se encontraron archivos sin duplicar.en: ", DIRECTORIO_ACTUAL)
        

def mostrar_archivos():
    n = 1
    directorio = os.listdir()                
    x = recorrer_duplicados(directorio,n)
    if x == 1:
        if six.PY3:
            input("Presiona <ENTER> para continuar")
            print("Abortando operación...")
            time.sleep(2)
            sys.exit(0)
    
def elegir_op(op):
        op = input("si - continuar | no - salir ")
        while op != ("Si" or op == "SI" or op == "si") or (op == "No" or op == "NO" or op == "no"):
            if op == "Si" or op == "SI" or op == "si":
                return eliminar_duplicados()
                break
            elif op == "No" or op == "NO" or op == "no":
                print("Hasta luego.")
                time.sleep(2)
                sys.exit(0)
            else:
                print("Opción incorrecta. \n")
                op = input("Continuar? si | no ")
    
def eliminar_duplicados():
    mostrar_archivos()
    op = ""
    #Filtrar por extensiones xlsx
    try:
        leer_libro = input("Elegí un documento: ")
        if "sin duplicar" in leer_libro:
            print("Error. No se puede elegir un archivo sin duplicar\n")
            return eliminar_duplicados() #vuelve al principio a modo de GoTo y evita que se duplique el archivo

        read_file = xlrd.open_workbook(leer_libro+'.xlsx')
        write_file = xlsxwriter.Workbook(leer_libro+' sin duplicar.xlsx')
        for hoja in read_file.sheets():
            filas = hoja.nrows        #devuelve el número de filas
            no_cols = hoja.ncols      #devuelve el número de columnas 
            nombre_hoja = hoja.name   #devuelve el nombre de las hojas del libro
            gen_sheets = write_file.add_worksheet(nombre_hoja) #escribe en cada nombre de hoja
            lista_solapas = []        
            cont = 0
            for fila in range(0, filas):                #Lee todas las filas 
                line_sublist = [hoja.cell(fila, col).value for col in range(0, no_cols)]
                if line_sublist not in lista_solapas:
                    lista_solapas.append(line_sublist)
                    for col in range(0, no_cols):       #Lee todas las columnas
                        gen_sheets.write(cont,col,line_sublist[col])
                    cont = cont + 1         
        write_file.close()
        print("Se borraron duplicados del archivo: ", leer_libro, " satisfactoriamente. ")
        print("Se creó el archivo " + leer_libro + " sin duplicar.xlsx" )
        elegir_op(op)
    except FileNotFoundError:
        print("ARCHIVO INEXISTENTE \n")
        elegir_op(op)
        
eliminar_duplicados()





    



