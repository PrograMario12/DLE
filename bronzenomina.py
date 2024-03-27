import csv

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
