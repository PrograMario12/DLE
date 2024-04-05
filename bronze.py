import csv
from datetime import datetime, timedelta
import pyodbc


def ConvertirStandardTimeEnHoras(segundos):
    horas = segundos / 3600 * 1.5
    return horas

def nomina():
    # Ruta del archivo de entrada y salida
    archivo_entrada = "assets/Nomina.csv"
    archivo_salida = "informacion_empleados.csv"
    tiempocomida = 0.5

    # Columnas a extraer
    columnas_extraer = ["# Empleado", "Nombre", "Departamento", "Linea", "Centro de Costos", "Turno"]

    dias = [["horario_LUN", "LUN E", "LUN S", "LUN Horas", "INC LUN"],
            ["horario_MAR", "MAR E", "MAR S", "MAR Horas", "INC MAR"],
            ["horario_MIE", "MIE E", "MIE S", "MIE Horas", "INC MIE"],
            ["horario_JUE", "JUE E", "JUE S", "JUE Horas", "INC JUE"],
            ["horario_VIE", "VIE E", "VIE S", "VIE Horas", "INC VIE"],
            ["horario_SAB", "SAB E", "SAB S", "SAB Horas", "INC SAB"],
            ["horario_DOM", "DOM E", "DOM S", "DOM Horas", "INC DOM"]]

    # Obtener la fecha actual
    fecha_actual = datetime.now().date()

    # Obtener el día de la semana (0: lunes, 1: martes, ..., 6: domingo)
    dia_semana = fecha_actual.weekday()

    informacion_empleados = []


    for i in range(7):
        print(i)
        # Calcular la diferencia de días para llegar al lunes
        dias_diferencia = dia_semana - i  # 1 representa el martes

        # Calcular la fecha del lunes
        fecha_registro = fecha_actual - timedelta(days=dias_diferencia)

        # Leer el archivo de entrada y extraer las columnas especificadas
        with open(archivo_entrada, "r") as file:
            reader = csv.DictReader(file)
            for row in reader:
                incidencias = row[dias[i][4]]
                horas = float(row[dias[i][3]]) if row[dias[i][3]] else 0

                departamento = row["Departamento"]
                arranque = 0.25 if "Inyección" in departamento else 0

                if incidencias:
                    horas = tiempocomida + arranque
                else: 
                    horas = 0

                HorasTotales = row[dias[i][3]]
                presencia = 0 if HorasTotales == 0 else (float(HorasTotales) - tiempocomida - arranque + horas)

                empleado = {columna: row[columna] for columna in columnas_extraer}
                empleado["Incidencia"] = incidencias
                empleado["Horas trabajadas"] = HorasTotales
                empleado["Tiempo de incidencias"] = horas
                empleado["Arranque"] = arranque
                empleado["Presencia"] = presencia
                empleado["Fecha"] = fecha_registro
                informacion_empleados.append(empleado)

                #print(empleado)

    # Escribir la información extraída en un nuevo archivo CSV
    columnas_extraer.append("Incidencia")
    columnas_extraer.append("Horas trabajadas")
    columnas_extraer.append("Tiempo de incidencias")
    columnas_extraer.append("Arranque")
    columnas_extraer.append("Presencia")
    columnas_extraer.append("Fecha")

    #Access

    # Establecer la conexión con la base de datos de Access
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\mariabar\OneDrive - Magna\Direct Labor Efficency\DLE.accdb;')

    # Crear un cursor para ejecutar consultas SQL
    cursor = conn.cursor()

    # query = "CREATE TABLE IF NOT EXISTS Nomina ([ID Empleado] TEXT, [Nombre] TEXT, [Departamento] TEXT, [Linea TEXT], [Centro De Costos] TEXT, [Turno] TEXT, [Incidencia] TEXT, [Horas Trabajadas] TEXT, [Tiempo de Incidencia] TEXT, [Arranque] TEXT, [Presencia] TEXT, [Fecha] TEXT)"

    # # Crear la tabla en la base de datos de Access (si no existe)
    # cursor.execute(query)

    # Insertar los datos en la tabla
    for empleado in informacion_empleados:
        cursor.execute('INSERT INTO Nomina (ID_Empleado, Nombre, Departamento, Linea, Centro_De_Costos, Turno, Incidencia, Horas_Trabajadas, Tiempo_De_Incidencia, Arranque, Presencia, Fecha) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', (empleado['# Empleado'], empleado['Nombre'], empleado['Departamento'], empleado['Linea'], empleado['Centro de Costos'], empleado['Turno'], empleado['Incidencia'], empleado['Horas trabajadas'], empleado['Tiempo de incidencias'], empleado['Arranque'], empleado['Presencia'], empleado['Fecha']))


    # Confirmar los cambios y cerrar la conexión
    conn.commit()
    conn.close()

    print("Información de empleados guardada en", archivo_salida)

########################################################### SAP ######################################
def sap():
    archivo_entrada_sap = "assets/SAP.csv"
    archivo_entrada_estandares = "assets/Estandares.csv"
    conexion_access = r"C:\Users\mariabar\OneDrive - Magna\Direct Labor Efficency\DLE.accdb;"

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
            local_entry_time = row["Local Entry Time"]
            fecha_hora = datetime.strptime(local_entry_time, "%d.%m.%Y %H:%M:%S")
            local_entry_time = fecha_hora.strftime("%d/%m/%Y")


            if movement_type in ["551", "131", "132", "552"]:
                if material in conteo_material:
                    conteo_material[material][movement_type] = conteo_material[material].get(movement_type, 0) + qty
                    #print(conteo_material)
                else:
                    conteo_material[material] = {movement_type: qty}

    #Obtener el total de QTY
    conteo_qty = {}

    for clave, valor in conteo_material.items():
        suma_qty = sum(valor.values())
        conteo_qty[clave] = suma_qty
        

    # Obtener la lista de todos los Movement Types
    movement_types = ["131", "132", "551", "552"]

    # Crear la conexión a la base de datos de Access
    conn = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + conexion_access)
    cursor = conn.cursor()

    # Insertar los datos en la tabla de Access
    for material, movement_counts in conteo_material.items():
        estandar = estandares.get(material, {"Dep": "", "Work center": "", "Standard Value": ""})
        dep = estandar["Dep"]
        work_center = estandar["Work center"]
        qtyfinal = str(conteo_qty.get(material, {"valor": ""}))
        if estandar["Standard Value"] != "":
            valor_sin_coma = estandar["Standard Value"].replace(',', '')  # Eliminar la coma del valor
            seh = ConvertirStandardTimeEnHoras(float(valor_sin_coma)) * float(qtyfinal)
            seh = round(seh, 2)
        else:
            seh = 0.0  # O cualquier otro valor predeterminado que desees asignar

        id = str(material) +  "" + str(local_entry_time).replace("/", "-")
        row = [id, dep, work_center, material, qtyfinal, seh, local_entry_time]
        for movement_type in movement_types:
            count = movement_counts.get(movement_type, 0)
            row.append(count)

        # Insertar la fila en la tabla de Access
        cursor.execute("INSERT INTO SAP (ID, Departament, Machine, Material, QTY_In_Unit_Of_Entry, Standard_earned_hours, Local_Time_Entry, Movement_type_131, Movement_type_132, Movement_type_551, Movement_type_552) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", row)

    # Guardar los cambios y cerrar la conexión
    conn.commit()
    conn.close()

    print("Conteo de repeticiones guardado en", archivo_salida)

sap()