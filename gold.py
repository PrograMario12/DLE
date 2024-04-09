from prettytable import PrettyTable
import csv
import pyodbc

def dlefinal():
    # Establecer la conexión con la base de datos de Access
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\mariabar\OneDrive - Magna\Direct Labor Efficency\DLE.accdb')

    # Crear un cursor para ejecutar consultas
    cursor = conn.cursor()

    # Ejecutar una consulta para obtener los datos que coincidan en la misma fecha
    consultaInyeccion = '''
    SELECT SUM(HG.EarnedHours) AS Total_EarnedHours, HP.Total_de_presencia, HP.Fecha_registro
    FROM HorasGanadas AS HG, HorasPagadas AS HP
    WHERE HG.Fecha = HP.Fecha_registro AND HP.Linea = 'Inyección' AND HG.Area = 'Inyección'
    GROUP BY HG.Area, HP.Total_de_presencia, HP.Fecha_registro
    '''

    consultaMetalizado = '''
    SELECT SUM(HG.EarnedHours) AS Total_EarnedHours, HP.Total_de_presencia, HP.Fecha_registro
    FROM HorasGanadas AS HG, HorasPagadas AS HP
    WHERE HG.Fecha = HP.Fecha_registro AND HP.Linea = 'Metalizado' AND HG.Area = 'Metalizado'
    GROUP BY HG.Area, HP.Total_de_presencia, HP.Fecha_registro
    '''

    consultaSubEnsambles = '''
    SELECT SUM(HG.EarnedHours) AS Total_EarnedHours, HP.Total_de_presencia, HP.Fecha_registro
    FROM HorasGanadas AS HG, HorasPagadas AS HP
    WHERE HG.Fecha = HP.Fecha_registro AND HP.Linea = 'SUBENSAMBLE' AND HG.Area = 'Subemsambles'
    GROUP BY HG.Area, HP.Total_de_presencia, HP.Fecha_registro
    '''

    consultaRefacciones = '''
    SELECT SUM(HG.EarnedHours) AS Total_EarnedHours, HP.Total_de_presencia, HP.Fecha_registro
    FROM HorasGanadas AS HG, HorasPagadas AS HP
    WHERE HG.Fecha = HP.Fecha_registro AND HP.Linea = 'REFACCIONES' AND HG.Area = 'REFACCIONES'
    GROUP BY HG.Area, HP.Total_de_presencia, HP.Fecha_registro
    '''

    consultaASM01 = '''
    SELECT SUM(HG.EarnedHours) AS Total_EarnedHours, HP.Total_de_presencia, HP.Fecha_registro
    FROM HorasGanadas AS HG, HorasPagadas AS HP
    WHERE HG.Fecha = HP.Fecha_registro AND HP.Linea = 'Metalizado' AND HG.Area = 'Metalizado'
    GROUP BY HG.Area, HP.Total_de_presencia, HP.Fecha_registro
    '''

    consultaasm06 = '''
    SELECT SUM(HG.EarnedHours) AS Total_EarnedHours, HP.Total_de_presencia, HP.Fecha_registro
    FROM HorasGanadas AS HG, HorasPagadas AS HP
    WHERE HG.Fecha = HP.Fecha_registro AND HP.Linea = 'Metalizado' AND HG.Area = 'Metalizado'
    GROUP BY HG.Area, HP.Total_de_presencia, HP.Fecha_registro
    '''

    cursor.execute(consultaInyeccion)

    # Obtener los resultados de la consulta
    resultados = cursor.fetchall()
    cursor.execute(consultaMetalizado)
    resultados.extend(cursor.fetchall())
    cursor.execute(consultaSubEnsambles)
    resultados.extend(cursor.fetchall())
    cursor.execute(consultaRefacciones)
    resultados.extend(cursor.fetchall())


    # Imprimir los resultados
    # for fila in resultados:
    #     print(fila)

    # Cerrar la conexión y el cursor
    cursor.close()
    conn.close()

    print(resultados)

dlefinal()