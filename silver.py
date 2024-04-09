import pyodbc
import csv

def resultadosrh():
    # Crear un diccionario para almacenar los totales por línea y área
    totales_linea = {}

    # Establecer la cadena de conexión a la base de datos de Access
    cadena_conexion = r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\mariabar\OneDrive - Magna\Direct Labor Efficency\DLE.accdb;"
    
    # Establecer la consulta SQL para obtener los datos
    consulta_sql = "SELECT Linea, Presencia, Fecha, Departamento FROM Nomina"

    # Crear la conexión a la base de datos de Access
    conn = pyodbc.connect(cadena_conexion)

    # Crear un cursor para ejecutar la consulta SQL
    cursor = conn.cursor()

    # Ejecutar la consulta SQL y obtener los datos
    cursor.execute(consulta_sql)
    datos = cursor.fetchall()

    # Diccionario para mapear líneas a áreas
    work_centers = {
        "injection" : ["MX_INY01", "MX_INY02", "MX_INY03", "MX_INY04", "MX_INY05", "MX_INY06", "MX_INY07", "MX_INY08", "MX_INY09", "MX_INY10", "MX_INY11", "MX_INY12", "MX_INY13", "MX_INY14", "MX_INY15", "MX_INY16", "MX_INY17", "MX_INY18", "MX_INY19", "MX_INY20", "MX_INY21", "MX_INY22", "MX_INY23", "MX_INY24", "MX_INY25", "MX_INY26", "MX_INY27", "MX_INY28", "MX_INY29", "MX_INY30", "MX_INY31"], 
        "metalizing" : ["MX-MET01", "MX-MET02", "MX-MET03", "MX-MET04", "MX-MET05", "MX-MET06", "MX-MET07", "MX-MET08"],
        "assembly" : ["MX_ASMSA", "MX_ASMRF", "MX_ASM01", "MX_ASM02", "MX_ASM03", "MX_ASM04", "MX_ASM05", "MX_ASM06", "MX_ASM07", "MX_ASM08", "MX_ASM09", "MX_ASM10", "MX_ASM11", "MX_ASM12", "MX_ASM13", "MX_ASM14", "MX_ASM15", "MX_ASM16", "MX_ASM17", "MX_ASM18", "MX_ASM19", "MX_ASM20", "MX_ASM21"]
    }

    for registro in datos:
        linea = registro.Linea
        if not linea:
            linea = registro.Departamento
            linea = linea.replace(" Directos", "")
        presencia = round(float(registro.Presencia), 2)
        fecha = registro.Fecha

        # Obtener el área correspondiente a la línea
        area = None
        for key, value in work_centers.items():
            if linea in value:
                area = key
                break

        # Verificar si la combinación de línea y área ya existe en el diccionario
        clave = (linea, area)
        if clave in totales_linea:
            # Si la combinación ya existe, sumar la presencia al total existente
            totales_linea[clave] += presencia
        else:
            # Si la combinación no existe, crear una nueva entrada en el diccionario
            totales_linea[clave] = presencia

    fecha = fecha.strftime("%d/%m/%Y")

    try:
        cursor.execute("CREATE TABLE HorasPagadas (Linea TEXT, Area TEXT, Total_de_presencia FLOAT, Fecha_registro DATETIME, Id TEXT)")
    except pyodbc.ProgrammingError as e:
        if e.args[0] == '42S01':
            print("La tabla 'HorasPagadas' ya existe.")
        else:
            print("Ocurrió un error:", e)

    # Insertar los datos en la tabla de Access
    for (linea, area), total in totales_linea.items():
        id = str(linea).replace(" ","") +  "" + str(fecha).replace("/", "")
        cursor.execute("INSERT INTO HorasPagadas (Linea, Area, Total_de_presencia, Fecha_registro, Id) VALUES (?, ?, ?, ?, ?)", linea, area, total, fecha, id)

    # Guardar los cambios y cerrar la conexión
    conn.commit()
    conn.close()
    print("El archivo totales_departamento.csv ha sido creado exitosamente")


def resultadossap():
    # Crear un diccionario para almacenar los totales por work center
    totales_work_center = {}

    # Establecer la cadena de conexión a la base de datos de Access
    cadena_conexion = r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\mariabar\OneDrive - Magna\Direct Labor Efficency\DLE.accdb;"
    
    # Establecer la consulta SQL para obtener los datos
    consulta_sql = "SELECT Machine, QTY_In_Unit_Of_Entry, Local_Time_Entry FROM SAP"

    # Crear la conexión a la base de datos de Access
    conn = pyodbc.connect(cadena_conexion)

    # Crear un cursor para ejecutar la consulta SQL
    cursor = conn.cursor()

    # Ejecutar la consulta SQL y obtener los datos
    cursor.execute(consulta_sql)
    datos = cursor.fetchall()

    # Diccionario para mapear máquinas a áreas
    work_centers = {
        "Inyección" : ["MX_INY01", "MX_INY02", "MX_INY03", "MX_INY04", "MX_INY05", "MX_INY06", "MX_INY07", "MX_INY08", "MX_INY09", "MX_INY10", "MX_INY11", "MX_INY12", "MX_INY13", "MX_INY14", "MX_INY15", "MX_INY16", "MX_INY17", "MX_INY18", "MX_INY19", "MX_INY20", "MX_INY21", "MX_INY22", "MX_INY23", "MX_INY24", "MX_INY25", "MX_INY26", "MX_INY27", "MX_INY28", "MX_INY29", "MX_INY30", "MX_INY31"], 
        "Metalizado" : ["MX-MET01", "MX-MET02", "MX-MET03", "MX-MET04", "MX-MET05", "MX-MET06", "MX-MET07", "MX-MET08"],
        "Subemsambles" : ["MX_ASMSA"],
        "Refacciones" : ["MX_ASMRF", "MX_ASM17","MX_ASM11", "MX_ASM12", "MX_ASM14", "MX_ASM15", "MX_ASM10"],
        "MX_ASM01" : ["MX_ASM01"],              
        "MX_ASM06" : ["MX_ASM06"],
        "MX_ASM07" : ["MX_ASM07", "MX_ASM03", "MX_ASM04",  "MX_ASM09"],
        "MX_ASM13-20" : ["MX_ASM13"],
        "MX_ASM16" : ["MX_ASM16"],
        "MX_ASM19" : [ "MX_ASM19"],  
        "MX_ASM21" : ["MX_ASM21"],          
        "Otras" : ["MX_ASM02",   "MX_ASM05",  "MX_ASM08",   "MX_ASM18","MX_ASM20"]
    }

    for registro in datos:
        machine = registro.Machine
        qty = round(registro.QTY_In_Unit_Of_Entry, 2)
        fecha = registro.Local_Time_Entry

        # Obtener el área correspondiente a la máquina
        area = None
        for key, value in work_centers.items():
            if machine in value:
                area = key
                break
        
        # Verificar si la combinación de machine y área ya existe en el diccionario
        clave = (machine, area)
        if clave in totales_work_center:
            # Si la combinación ya existe, sumar la cantidad al total existente
            totales_work_center[clave] += qty
        else:
            # Si la combinación no existe, crear una nueva entrada en el diccionario
            totales_work_center[clave] = qty

    fecha = fecha.strftime("%d/%m/%Y")

    try:
        cursor.execute("CREATE TABLE HorasGanadas (Machine TEXT, Area TEXT, EarnedHours FLOAT, Id TEXT, Fecha TEXT)")
    except pyodbc.ProgrammingError as e:
        if e.args[0] == '42S01':
            print("La tabla 'HorasGanadas' ya existe.")
        else:
            print("Ocurrió un error:", e)

    # Insertar los datos en la tabla de Access
    for (machine, area), qty_total in totales_work_center.items():
        id = str(machine).replace(" ","") +  "" + str(fecha).replace("/", "")
        cursor.execute("INSERT INTO HorasGanadas (Machine, Area, EarnedHours, Id, Fecha) VALUES (?, ?, ?, ?, ?)", machine, area, qty_total, id, fecha)

    # Guardar los cambios y cerrar la conexión
    conn.commit()
    conn.close()


    print("El archivo totales_work_center.csv ha sido creado exitosamente.")


#resultadosrh()
resultadossap()