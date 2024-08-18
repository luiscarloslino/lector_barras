import os
from datetime import datetime
import pandas as pd
import sys

# Redirigir stderr a un archivo para evitar ver las advertencias en la consola
sys.stderr = open('error_log.txt', 'w')

# Definir el diccionario con los datos
diccionario = {
    "29593447": [1, "29593447", "993481701", "BALLON CANO, GHINA", "PROMOTORA", "PROMOTORA"],
    "40562059": [2, "40562059", "983110909", "AGUIRRE CUTIPA, HENRY YOEL", "MAXIMO", "PRESTACION"],
    "29366764": [3, "29366764", "958273066", "AGUIRRE GUZMAN DE CHIRINOS, IVIS LOURDES", "MAXIMO", "FISCALIZACIÓN"],
    "45655378": [4, "45655378", "958216619", "ALFARO ORIHUELA, JORGE ELVIS", "MAXIMO", "ECONOMÍA"],
    "44690762": [5, "44690762", "970777414", "ALLER ALLER, CRISTIAN", "WENDY", "TRABAJO SOCIAL"],
    "72322452": [6, "72322452", "922901029", "ALVAREZ GONZALES, FRANK REY", "ROCIO", "TRABAJO SOCIAL"],
    "80498608": [7, "80498608", "913242011", "ALVAREZ HANAMPA, LORENZO", "MARILU", "ECONOMÍA"],
    "47996952": [8, "47996952", "929966788", "ALVAREZ HUAYHUA, WILNOR IVAN", "MARILU", "ADMINISTRACIÓN"],
    "29344434": [9, "29344434", "982381763", "ALVAREZ IBARCENA, AURELIA ANANI", "W. MOLINA", "PRESTACION"],
    "70129774": [10, "70129774", "955703353", "ALVAREZ MENACHO, JOHAN ANDREE", "MARILU", "PRESTACION"],
    "47620916": [11, "47620916", "992543528", "ALVAREZ PAYAHUANCA, ALEX LEONCIO", "MARILU", "PRESTACION"],
    "29664230": [12, "29664230", "925134147", "ANCO CHARCA, MATILDE CECILIA", "MAXIMO", "ABASTECIMIENTOS "],
}

# Diccionario para almacenar las horas de entrada y salida
registro_asistencia = {}

# Diccionario para contar las veces que se ingresa un DNI
contador_dnis = {}

# Contador de registros autoincrementable
contador_registros = 0

# Función para buscar y mostrar la información
def buscar_persona_por_dni(dni):
    return buscar_persona(dni, 1)

def buscar_persona_por_nombre(nombre):
    return buscar_persona(nombre, 3)

def buscar_persona_por_numero(numero):
    return buscar_persona(numero, 0)

def buscar_persona(valor, campo):
    global contador_registros
    current_time = datetime.now().strftime("%I:%M:%S %p")  # Obtener la hora actual

    encontrado = False
    for dni, informacion in diccionario.items():
        if str(informacion[campo]) == valor:
            encontrado = True
            nombre_completo = informacion[3]

            # Actualizar el contador de DNI
            if dni in contador_dnis:
                contador_dnis[dni] += 1
                print(f"El DNI {dni} ya ha sido registrado {contador_dnis[dni]} veces.")
            else:
                contador_dnis[dni] = 1

            # Comprobar si ya existe una entrada para el DNI en el registro
            if dni not in registro_asistencia:
                registro_asistencia[dni] = {
                    "Nombre": nombre_completo,
                    "Entrada": current_time,
                    "Salida": "",
                    "N°": informacion[0],
                    "Celular": informacion[2],
                    "Lider": informacion[4],
                    "Cargo": informacion[5],
                    "Puesto": contador_registros + 1
                }
                contador_registros += 1
            else:
                if registro_asistencia[dni]["Salida"] == "":
                    registro_asistencia[dni]["Salida"] = current_time
                    return
                else:
                    return

            # Mostrar la información en la consola
            print(f"{contador_registros} ▶ {current_time}")
            print(f"N°: {informacion[0]}")
            print(f"DNI: {informacion[1]}")
            print(f"Lider: {informacion[4]}")
            print(f"Nombre: {informacion[3]}")
            print(f"Celular: {informacion[2]}")
            print(f"Cargo: {informacion[5]}")
            print("------------------------------------------------------------")
            return True  # Retorna True indicando que se encontró y mostró la información

    if not encontrado:
        print(f"{current_time} ⇒ No se encontró información para el valor {valor}.")
        print("------------------------------------------------------------")
        return False  # Retorna False indicando que no se encontró la información

# Función para guardar los registros en un archivo Excel
def guardar_registros_en_excel():
    registros = []
    for dni, informacion in diccionario.items():
        entrada = registro_asistencia[dni]["Entrada"] if dni in registro_asistencia else ""
        pago_entrada = 2 if dni in registro_asistencia else ""
        puesto = registro_asistencia[dni]["Puesto"] if dni in registro_asistencia else ""
        if entrada:
            entrada_time = datetime.strptime(entrada, "%I:%M:%S %p")
            hora_limite = datetime.strptime("20:15", "%H:%M")
            puntualidad = "ASISTIO" if entrada_time <= hora_limite else "TARDANZA"
        else:
            puntualidad = "INASISTENCIA"

        registros.append({
            "Puesto": puesto,
            "N°": informacion[0],
            "DNI": dni,
            "Apellidos y nombres": informacion[3],
            "Lider": informacion[4],
            "Cargo": informacion[5],
            "Hora de entrada": entrada,
            "Pago Entrada": pago_entrada,
            "Puntualidad": puntualidad
        })
    df = pd.DataFrame(registros)
    fecha_actual = datetime.now().strftime("%d-%m-%Y")
    nombre_archivo = f"REG_ENTRADA_{fecha_actual}.xlsx"
    df.to_excel(nombre_archivo, index=False, columns=["Puesto", "N°", "DNI", "Apellidos y nombres", "Lider", "Cargo", "Hora de entrada", "Pago Entrada", "Puntualidad"])
    print(f"Registros guardados en {nombre_archivo}")

# Función para manejar el ingreso del DNI desde la terminal
def ingresar_dni_terminal():
    while True:
        dni = input("Ingrese el DNI (o 'salir' para terminar): ")
        if dni.lower() == 'salir':
            break
        buscar_persona_por_dni(dni)

# Ejecutar el ingreso de DNI en el hilo principal
ingresar_dni_terminal()

# Guardar los registros en un archivo Excel al finalizar
guardar_registros_en_excel()
