# Importar las bibliotecas necesarias
import pandas as pd
import numpy_financial as npf
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
import time 
from datetime import datetime
from docx import Document
import os
from datetime import datetime
import comtypes.client
import shutil,sys

# importamos las funciones necesarias
from funciones import *

#print("Iniciando")
#nombre_cliente = '  '
def main(nombre_cliente,plazo,fecha,Numero):
    directorio_actual = os.path.dirname(os.path.abspath(__file__))
    #print("primero",directorio_actual)
    # Cargamos los archivos necesarios
    Reporte_ventas = os.path.join(directorio_actual,"Reporte de Ventas.xlsx")
    df_ventas = pd.read_excel(Reporte_ventas)
    #df_ventas = pd.read_excel('C:/Users/alber/Downloads/Reporte de Ventas.xlsx')
    Calendario = os.path.join(directorio_actual,"db_calendario.xlsx")
    df_calendario = pd.read_excel(Calendario)
    #df_calendario = pd.read_excel("C:/Users/alber/Downloads/db_calendario.xlsx")

    # Guardamos en una variable la fecha del dia de hoy
    fecha_hoy = pd.to_datetime('today').date()
    #print("PRIMERA MARCA")

    #-----------------Ejecutamos las funciones-----------------#
    numeros_credito = df_ventas[(df_ventas['Nombre del Cliente'] == nombre_cliente) & (df_ventas['EstatusCredito'] == 'Vigente')]['NumeroCredito'].tolist()
    

    if len(numeros_credito) < 1:
        print()
        #print(f"El cliente {nombre_cliente} no tiene creditos vigentes")
    else:
        #print((numeros_credito))
        # Medimos el tiempo de ejecución
        start_time = time.time()
        #print("SEGUNDA MARCA")

        # Creamos una carpeta con el nombre del cliente
        nombre_cliente = nombre_cliente.rstrip()  # Elimina los espacios en blanco al final
        ruta_carpeta = os.path.join(directorio_actual,nombre_cliente)
        #ruta_carpeta = r'C:\Users\alber\Downloads\{}'.format(nombre_cliente)
        ruta_zip = os.path.join(directorio_actual, nombre_cliente + '.zip')
        
        #En caso de existir una carpeta la elimina
        if os.path.exists(ruta_carpeta):
            #print("ya existe")
            shutil.rmtree(ruta_carpeta)  # Elimina la carpeta y todo su contenido
        else:
            #print("Carpeta creada")
            os.makedirs(ruta_carpeta, exist_ok=True)
        #En caso de existir un archivo ZIP lo elimina
        if os.path.exists(ruta_zip):
            #print(f"ZIP de {nombre_cliente} ya existe")
            os.remove(ruta_zip)  # Elimina la carpeta y todo su contenido
        else:
            #print("No existe un ZIP de este cliente")
            print()

        #print("TERCERA MARCA")
        # Iteramos sobre los números de crédito
        for numero_credito in numeros_credito:
            #print("CUARTA MARCA")
            dire = os.path.join(directorio_actual,"formatoCartaReestructuracion - copia.docx")
            doc = Document(dire)
            #print("CUARTA MARCA 1/2")
            #Convertimos numero_credito a string
            numero_credito = str(numero_credito)
            informacion = informacion_credito(numero_credito, df_ventas)
            df_pagos, df_movimientos = generacion_dataframes(numero_credito)
            df_amortizacion = tabla_amortizacion(informacion)
            pago_amortizacion = df_amortizacion['Pago'].sum()
            importe_movimientos = df_movimientos['Importe'].sum()
            pago = informacion['Pago']
            ultimo_indice = encontrar_ultimo_indice_menor_que_pago(df_pagos,pago)
            validacion, fecha_movimiento = comparar_fechas(df_pagos, df_movimientos, ultimo_indice)
            quincenas_pago = obtener_numero_quincenas(fecha_movimiento)
            quincenas_corte = obtener_numero_quincenas(pd.to_datetime(df_calendario['Fecha de corte'].iloc[0]))
            porcentaje_deuda = ((importe_movimientos + pago) / pago_amortizacion)
            #print("QUINTA MARCA")
            if quincenas_pago is not None:
                diferencia_quincena = quincenas_pago - quincenas_corte
            else:
                diferencia_quincena = None

            if validacion == False:
                #print('Las fechas de pago y movimientos no coinciden')
                #print('Es un credito con movimientos irregulares')
                # Crear un writer de Excel para cada número de crédito
                #print("quinta 1/2")
                nombre_archivo_excel = os.path.join(directorio_actual,f'credito_{numero_credito}.xlsx')
                #print("quinta 3/4")
                #nombre_archivo_excel = f'C:\\Users\\alber\\Downloads\\credito_{numero_credito}.xlsx'
                with pd.ExcelWriter(nombre_archivo_excel) as writer:
                # Guardar cada dataframe en una pestaña diferente
                    df_pagos.to_excel(writer, sheet_name='df_pagos')
                    df_movimientos.to_excel(writer, sheet_name='df_movimientos')
                    df_amortizacion.to_excel(writer, sheet_name='df_amortizacion')
                    
                #movemos el archivo creado a la carpeta del cliente que se almaceno en la variable ruta_carpeta
                os.rename(nombre_archivo_excel, os.path.join(ruta_carpeta, f'credito_{numero_credito}.xlsx'))
                continue
            #print("salta")
            # Realizamos los siguientes pasos si importe_movimientos es mayor o igual al 40% de pago_amortizacion
            if porcentaje_deuda >= 0.35 and validacion == True:
                #print("123")
                df_calendario = informacion_calendario(df_calendario, informacion)
                df_cruzada = generar_tabla_cruzada(df_amortizacion, df_pagos, df_calendario, nombre_cliente, numero_credito)
                #print("1")
                # Rellenamos la carta
                rellenado_carta(doc, fecha_hoy, df_cruzada, informacion['idSolicitud'],nombre_cliente,fecha,plazo,Numero)
                #print("2")
                # Generamos el PDF
                generar_pdf(numero_credito)
                #print("SEXTA MARCA")
                # Eliminamos el archivo de word
                plantilla = os.path.join(directorio_actual,"plantilla.docx")
                os.remove(plantilla)
                #os.remove(r'C:\Users\alber\Downloads\plantilla.docx')
                #print("SEPTIMA MARCA")
                nombre_archivo_excel = os.path.join(directorio_actual,f"credito_{numero_credito}.xlsx")
                #nombre_archivo_excel = f'C:\\Users\\alber\\Downloads\\credito_{numero_credito}.xlsx'
                with pd.ExcelWriter(nombre_archivo_excel) as writer:
                    # Guardar cada dataframe en una pestaña diferente
                    df_pagos.to_excel(writer, sheet_name='df_pagos')
                    df_movimientos.to_excel(writer, sheet_name='df_movimientos')
                    df_amortizacion.to_excel(writer, sheet_name='df_amortizacion')

                # Verificamos si la carpeta del cliente existe, si no, la creamos
                if not os.path.exists(ruta_carpeta):
                    os.makedirs(ruta_carpeta)
                #print("OCTAVA MARCA")
                # Movemos el archivo de pdf creado a la nueva carpeta del cliente
                os.rename(os.path.join(directorio_actual,f"cartaReestructura_{numero_credito}.pdf") , os.path.join(ruta_carpeta, f'cartaReestructura_{numero_credito}.pdf'))
                #os.rename(r'C:\Users\alber\Downloads\cartaReestructura_{}.pdf'.format(numero_credito), os.path.join(ruta_carpeta, f'cartaReestructura_{numero_credito}.pdf'))
                #print("NOVENA MARCA")
                # Movemos los archivos creados a la carpeta del cliente
                os.rename(nombre_archivo_excel, os.path.join(ruta_carpeta, f'credito_{numero_credito}.xlsx'))
                #print("Archivo creado")
                #print("DECIMA MARCA")
                continue

            else:
                #print("septimo 1/2")
                #print('El importe de movimientos no es mayor o igual al 40% del pago amortizable')
                #print(f'No se generará la carta de reestructura del credito {numero_credito}')
                #print('El valor de validacion es: ', validacion)
                #print('El valor de porcentaje de deuda es: ', porcentaje_deuda)
                nombre_archivo_excel = os.path.join(directorio_actual,f"credito_{numero_credito}.xlsx")
                #nombre_archivo_excel = f'C:\\Users\\alber\\Downloads\\credito_{numero_credito}.xlsx'
                with pd.ExcelWriter(nombre_archivo_excel) as writer:
                # Guardar cada dataframe en una pestaña diferente
                    df_pagos.to_excel(writer, sheet_name='df_pagos')
                    df_movimientos.to_excel(writer, sheet_name='df_movimientos')
                    df_amortizacion.to_excel(writer, sheet_name='df_amortizacion')
                #("octavo")
                # Movemos el archivo creado a la carpeta del cliente
                os.rename(nombre_archivo_excel, os.path.join(ruta_carpeta, f'credito_{numero_credito}.xlsx'))

                continue

        # Convertimos la carpeta del cliente en un archivo ZIP
        ruta_del_archivo_generado = shutil.make_archive(ruta_carpeta, 'zip', ruta_carpeta)

        print("RUTA_ARCHIVO: " + ruta_del_archivo_generado)

        # Eliminamos la carpeta del cliente que no es zip
        shutil.rmtree(ruta_carpeta)

        # Medimos el tiempo de ejecución
        end_time = time.time()
        execution_time = end_time - start_time
        #print(f"\nTiempo de ejecución: {execution_time} segundos")

if __name__ == "__main__":
     # Verifica si se proporciona el argumento del nombre del cliente
    if len(sys.argv) != 5:
        #print("Uso: CartaReestructura.py <nombre_cliente>")
        print("Error")
        sys.exit(1)

    # Obtiene el nombre del cliente del primer argumento
    nombre_cliente = sys.argv[1]
    plazo = sys.argv[2]
    fecha = sys.argv[3]
    Numero = sys.argv[4]

    # Llama a la función principal con el nombre del cliente
    main(nombre_cliente,plazo,fecha,Numero)