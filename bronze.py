import csv

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

def sap():
    import csv

    archivo_entrada = "assets/SAP.csv"
    archivo_salida = "conteo_material.csv"

    # Diccionario para almacenar el conteo de repeticiones
    conteo_material = {}

    # Leer el archivo de entrada y contar las repeticiones
    with open(archivo_entrada, "r") as file:
        reader = csv.DictReader(file)
        for row in reader:
            movement_type = row["Movement Type"]
            material = row["Material"]

            if movement_type in ["551", "131", "132", "552"]:
                if material in conteo_material:
                    conteo_material[material][movement_type] = conteo_material[material].get(movement_type, 0) + 1
                else:
                    conteo_material[material] = {movement_type: 1}

    # Obtener la lista de todos los Movement Types
    movement_types = ["131", "132", "551", "552"]

    # Escribir el conteo de repeticiones en un nuevo archivo CSV
    with open(archivo_salida, "w", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["Material"] + movement_types)
        for material, movement_counts in conteo_material.items():
            row = [material]
            for movement_type in movement_types:
                count = movement_counts.get(movement_type, 0)
                row.append(count)
            writer.writerow(row)

    print("Conteo de repeticiones guardado en", archivo_salida)

sap()