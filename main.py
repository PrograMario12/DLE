import tkinter as tk
import bronze, silver, gold

def calcular_dle():
    # ventanasaludos = tk.Tk()
    # ventanasaludos.title("Resultados DLE")    

    # # Crear el título
    # titulo = tk.Label(ventanasaludos, text="Proceso ejecuntandose", font=("Arial", 12))
    # titulo.pack()

    bronze.nomina()
    bronze.sap()

    silver.resultadossap()
    silver.resultadosrh()

    gold.creartabla()
    pass

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Cálculo del DLE")

# Establecer el tamaño de la ventana
ventana.geometry("500x500")

# Obtener el ancho y alto de la pantalla
ancho_pantalla = ventana.winfo_screenwidth()
alto_pantalla = ventana.winfo_screenheight()

# Calcular las coordenadas para centrar la ventana
posicion_x = int(ancho_pantalla / 2 - 250)
posicion_y = int(alto_pantalla / 2 - 250)

# Centrar la ventana en la pantalla
ventana.geometry("+{}+{}".format(posicion_x, posicion_y))

# Crear el título
titulo = tk.Label(ventana, text="Cálculo de DLE", font=("Arial", 30))
titulo.pack()

# Crear el botón
boton_ejecutar = tk.Button(ventana, text="Ejecutar", command=calcular_dle, width=30, height=2)
boton_ejecutar.pack()

# Iniciar el bucle de eventos
ventana.mainloop()