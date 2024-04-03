import csv

def resultadosrh():
    # Crear un diccionario para almacenar los totales por departamento
    totales_departamento = {}

    # Leer el archivo CSV
    with open('informacion_empleados.csv', 'r') as archivo_csv:
        lector_csv = csv.DictReader(archivo_csv)
        
        # Iterar sobre cada registro del archivo CSV
        for registro in lector_csv:
            departamento = registro['Departamento']
            presencia = round(float(registro['Presencia']),2)
            
            # Verificar si el departamento ya existe en el diccionario
            if departamento in totales_departamento:
                # Si el departamento ya existe, sumar la presencia al total existente
                totales_departamento[departamento] += presencia
            else:
                # Si el departamento no existe, crear una nueva entrada en el diccionario
                totales_departamento[departamento] = presencia

    # Imprimir los totales por departamento
    # for departamento, total in totales_departamento.items():
    #     print(f"Departamento: {departamento} - Total de presencia: {total}")
                
    with open('silver/totales_departamento.csv', 'w', newline='') as archivo_salida:
        campos = ['Departamento', 'Total de presencia']
        escritor_csv = csv.DictWriter(archivo_salida, fieldnames=campos)

        #Escribir la fila de encabezado
        escritor_csv.writeheader

        #Escribir los datos de cada departamento
        for departamento, total in totales_departamento.items():
            escritor_csv.writerow({'Departamento': departamento, 'Total de presencia': total})

    print("El archivo totales_departamento.csv ha sido creado exitosamente")

def resultadossap():
    # Crear un diccionario para almacenar los totales por work center
    totales_work_center = {}

    # Leer el archivo CSV
    with open('conteo_material.csv', 'r') as archivo_csv:
        lector_csv = csv.DictReader(archivo_csv)
        
        # Iterar sobre cada registro del archivo CSV
        for registro in lector_csv:
            work_center = registro['Work center']
            qty = round(float(registro['QTY']),2)
            
            # Verificar si el work center ya existe en el diccionario
            if work_center in totales_work_center:
                # Si el work center ya existe, sumar la cantidad al total existente
                totales_work_center[work_center] += qty
            else:
                # Si el work center no existe, crear una nueva entrada en el diccionario
                totales_work_center[work_center] = qty

    # Guardar los totales por work center en un archivo CSV
    with open('silver/totales_work_center.csv', 'w', newline='') as archivo_salida:
        campos = ['Work center', 'Total de QTY']
        escritor_csv = csv.DictWriter(archivo_salida, fieldnames=campos)
        
        # Escribir la fila de encabezado
        escritor_csv.writeheader()
        
        # Escribir los datos de cada work center
        for work_center, total in totales_work_center.items():
            escritor_csv.writerow({'Work center': work_center, 'Total de QTY': total})

    print("El archivo totales_work_center.csv ha sido creado exitosamente.")


#resultadosrh()
#resultadossap()