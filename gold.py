from prettytable import PrettyTable
import csv

def creartabla():
    # Lista de work centers a considerar
    work_centers = {"injection" : ["MX_INY01", "MX_INY02", "MX_INY03", "MX_INY04", "MX_INY05", "MX_INY06", "MX_INY07", "MX_INY08", "MX_INY09", "MX_INY10", "MX_INY11", "MX_INY12", "MX_INY13", "MX_INY14", "MX_INY15", "MX_INY16", "MX_INY17", "MX_INY18", "MX_INY19", "MX_INY20", "MX_INY21", "MX_INY22", "MX_INY23", "MX_INY24", "MX_INY25", "MX_INY26", "MX_INY27", "MX_INY28", "MX_INY29", "MX_INY30", "MX_INY31"], 
    "metalizing" : ["MX-MET01", "MX-MET02", "MX-MET03", "MX-MET04", "MX-MET05", "MX-MET06", "MX-MET07", "MX-MET08"],
    "assembly" : ["MX_ASMSA", "MX_ASMRF", "MX_ASM01", "MX_ASM02", "MX_ASM03", "MX_ASM04", "MX_ASM05", "MX_ASM06", "MX_ASM07", "MX_ASM08", "MX_ASM09", "MX_ASM10", "MX_ASM11", "MX_ASM12", "MX_ASM13", "MX_ASM14", "MX_ASM15", "MX_ASM16", "MX_ASM17", "MX_ASM18", "MX_ASM19", "MX_ASM20", "MX_ASM21"]}

    suma_qty = {}

    # Inicializar las claves del diccionario suma_qty
    for area in work_centers:
        suma_qty[area] = 0

    # Leer el archivo CSV y sumar los valores
    with open('silver/totales_work_center.csv', 'r') as archivo_csv:
        lector_csv = csv.DictReader(archivo_csv)
        
        for registro in lector_csv:
            work_center = registro['Work center']
            qty = float(registro['Total de QTY'])
            
            # Verificar si el work center está en el diccionario suma_qty
            if work_center in work_centers["injection"]:
                suma_qty["injection"] += qty
            elif work_center in work_centers["metalizing"]:
                suma_qty["metalizing"] += qty
            elif work_center in work_centers["assembly"]:
                suma_qty["assembly"] += qty

    # Mostrar los resultados
    # for area, total in suma_qty.items():
    #     print(f"Total de QTY para el área {area}: {total}")
                
    # Retornar las horas pagadas
    horas_pagadas = {}

    # Leer el archivo CSV
    with open('silver/totales_departamento.csv', 'r') as archivo_csv:
        lector_csv = csv.reader(archivo_csv)
        
        # Iterar sobre cada registro del archivo CSV
        for registro in lector_csv:
            columna1 = registro[0]
            columna2 = registro[1]
            
            # Verificar si el valor de la columna1 está en la lista de valores buscados
            if columna1 in ["Inyección Directos", "Ensamble Directos", "Metalizado Directos"]:
                # Guardar el valor de la columna2 en el diccionario
                if columna1 == "Inyección Directos":
                    horas_pagadas["injection"] = columna2
                elif columna1 == "Ensamble Directos":
                    horas_pagadas["assembly"] = columna2
                else:
                    horas_pagadas["metalizing"] = columna2

    # Imprimir los valores encontrados
    #for clave, valor in horas_pagadas.items():
        #print(f"{clave}: {valor}")          

    # Crear una instancia de la tabla
    tabla = PrettyTable()
    tabla.field_names = ["Área", "Earned hours", "Paid hours", "DLE"]

    # Agregar filas a la tabla
    for area, valor in suma_qty.items():
        horas_pagadas_valor = horas_pagadas.get(area, 0)
        horas_pagadas_valor = float(horas_pagadas_valor)
        dle = round((valor / horas_pagadas_valor) * 100, 2)
        tabla.add_row([area, valor, horas_pagadas_valor, dle])

    # Imprimir la tabla
    print(tabla)