def calcular_dle(produccion_real, horas_trabajadas, produccion_estandar, horas_estandar):
    dle = (produccion_real * horas_estandar) / (produccion_estandar * horas_trabajadas)
    return dle

# Ejemplo de uso
produccion_real = 1000
horas_trabajadas = 160
produccion_estandar = 1200
horas_estandar = 160

dle = calcular_dle(produccion_real, horas_trabajadas, produccion_estandar, horas_estandar)
print("Eficiencia de Mano de Obra Directa:", dle)
