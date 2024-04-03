import csv

def ConvertirStandardTimeEnHoras(segundos):
    horas = segundos / 3600 * 1.5
    return horas

def nomina():
    # Ruta del archivo de entrada y salida
    archivo_entrada = "assets/Nomina.csv"
    archivo_salida = "informacion_empleados.csv"
    tiempocomida = 0.5

    # Columnas a extraer
    columnas_extraer = ["# Empleado", "Nombre", "Departamento", "Linea", "Centro de Costos", "Turno", "horario_MAR", "MAR E", "MAR S", "MAR Horas", "INC MAR"]

    # Leer el archivo de entrada y extraer las columnas especificadas
    informacion_empleados = []
    with open(archivo_entrada, "r") as file:
        reader = csv.DictReader(file)
        for row in reader:
            incidencias = row["INC MAR"]
            mar_horas = float(row["MAR Horas"]) if row["MAR Horas"] else 0.0

            departamento = row["Departamento"]
            arranque = 0.25 if "Inyección" in departamento else 0

            if incidencias:
                mar_horas += tiempocomida + arranque
            else: 
                mar_horas = 0

            HorasTotales = row["MAR Horas"]
            presencia = 0 if HorasTotales == 0 else (float(HorasTotales) - tiempocomida - arranque + mar_horas)

            empleado = {columna: row[columna] for columna in columnas_extraer}
            empleado["Incidencias"] = mar_horas
            empleado["Arranque"] = arranque
            empleado["Presencia"] = presencia
            informacion_empleados.append(empleado)

    # Escribir la información extraída en un nuevo archivo CSV
    columnas_extraer.append("Incidencias")
    columnas_extraer.append("Arranque")
    columnas_extraer.append("Presencia")
    with open(archivo_salida, "w", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=columnas_extraer)
        writer.writeheader()
        writer.writerows(informacion_empleados)

    print("Información de empleados guardada en", archivo_salida)




######## SAP ##########
def sap():
    archivo_entrada_sap = "assets/SAP.csv"
    archivo_entrada_estandares = "assets/Estandares.csv"
    archivo_salida = "conteo_material.csv"

    #Diccionario para almacenar los estándares de cada material
    estandares = {}

    # Leer el archivo de estándares y almacenar los datos
    with open(archivo_entrada_estandares, "r") as file:
        reader = csv.DictReader(file)
        for row in reader:
            material = str(row["Material"])
            area = row["AREA"]
            work_center = row["Work center"]
            standard_value = row["Standard Value"]
            estandares[material] = {"Dep": area, "Work center": work_center, "Standard Value": standard_value}


    # Diccionario para almacenar el conteo de repeticiones
    conteo_material = {}

    # Leer el archivo de entrada y contar las repeticiones
    with open(archivo_entrada_sap, "r") as file:
        reader = csv.DictReader(file)
        for row in reader:
            movement_type = row["Movement Type"]
            material = str(row["Material"])
            qty = float(row["Qty in unit of entry"])

            if movement_type in ["551", "131", "132", "552"]:
                if material in conteo_material:
                    conteo_material[material][movement_type] = conteo_material[material].get(movement_type, 0) + qty
                    #print(conteo_material)
                else:
                    conteo_material[material] = {movement_type: qty}
            
            #print(conteo_material)

    #Obtener el total de QTY
    conteo_qty = {}

    for clave, valor in conteo_material.items():
        suma_qty = sum(valor.values())
        conteo_qty[clave] = suma_qty

    #print(conteo_qty, "\n")
        

    # Obtener la lista de todos los Movement Types
    movement_types = ["131", "132", "551", "552"]

    # Escribir el conteo de repeticiones en un nuevo archivo CSV
    with open(archivo_salida, "w", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["Dep", "Work center", "Material", "QTY", "Standard Earned Hours"] + movement_types)
        for material, movement_counts in conteo_material.items():
            estandar = estandares.get(material, {"Dep": "", "Work center": "", "Standard Value": ""})
            #print(estandar)
            dep = estandar["Dep"]
            work_center = estandar["Work center"]
            qtyfinal = str(conteo_qty.get(material, {"valor": ""}))
            if estandar["Standard Value"] != "":
                valor_sin_coma = estandar["Standard Value"].replace(',', '')  # Eliminar la coma del valor
                seh = ConvertirStandardTimeEnHoras(float(valor_sin_coma)) * float(qtyfinal)
                seh = round(seh,2)
            else:
                seh = 0.0  # O cualquier otro valor predeterminado que desees asignar

            row = [dep, work_center, material, qtyfinal, seh]
            #print (row)
            for movement_type in movement_types:
                count = movement_counts.get(movement_type, 0)
                row.append(count)
            writer.writerow(row)

    print("Conteo de repeticiones guardado en", archivo_salida)