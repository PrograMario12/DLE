import csv

# Ruta del archivo de entrada y salida
archivo_entrada = "assets/Nomina.csv"
archivo_salida = "informacion_empleados.csv"

# Columnas a extraer
columnas_extraer = ["# Empleado", "Nombre", "Departamento", "Linea",  "Centro de Costos", "Turno", "horario_MAR", "MAR E", "MAR S", "MAR Horas", "INC MAR"]

# Leer el archivo de entrada y extraer las columnas especificadas
informacion_empleados = []
with open(archivo_entrada, "r") as file:
    reader = csv.DictReader(file)
    for row in reader:
        empleado = {columna: row[columna] for columna in columnas_extraer}
        informacion_empleados.append(empleado)

# Escribir la información extraída en un nuevo archivo CSV
with open(archivo_salida, "w", newline="") as file:
    writer = csv.DictWriter(file, fieldnames=columnas_extraer)
    writer.writeheader()
    writer.writerows(informacion_empleados)

print("Información de empleados guardada en", archivo_salida)
