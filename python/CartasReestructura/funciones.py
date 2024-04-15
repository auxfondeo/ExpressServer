import pandas as pd, shutil,time,os
import numpy_financial as npf
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from datetime import datetime
from docx import Document
from datetime import datetime
import comtypes.client

directorio_actual = os.path.dirname(os.path.abspath(__file__))
#print("Segundo",directorio_actual)

# Creamos una funcion para obtener la información de un crédito por su numero de credito
def informacion_credito(numero_credito, df_ventas):


    # Filtramos el df por el numero de credito
    df_credito = df_ventas[df_ventas['NumeroCredito'] == int(numero_credito)]

    # Verificamos si se encontró el crédito
    if df_credito.empty:
        return None

    # Obtenemos los valores de las columnas relevantes
    idSolicitud = df_credito['IdSolicitud'].values[0]
    monto_credito = df_credito['MontoCredito'].values[0]
    plazo = df_credito['Plazo'].values[0]
    pago = df_credito['Pago'].values[0]
    tasa_ordinaria = df_credito['TasaOrdinaria'].values[0] / 100
    institucion = df_credito['Institucion'].values[0]

    # Creamos un diccionario con la información del crédito
    informacion = {
        'idSolicitud': idSolicitud,
        'NumeroCredito': numero_credito,
        'MontoCredito': monto_credito,
        'Plazo': plazo,
        'Pago': pago,
        'TasaOrdinaria': tasa_ordinaria,
        'Institucion': institucion
    }

    # Devolvemos la información del crédito
    return informacion

# Creamos una funcion para crear cartas de reestructura
def generacion_dataframes(numero_credito):
    #print(f"Empezando con el credito: {numero_credito}")
    """
    Genera dos DataFrames a partir de un número de crédito.

    Parámetros:
    - numero_credito: El número de crédito del cual se generarán los DataFrames.

    Retorna:
    - df_pagos: DataFrame que contiene los pagos relacionados al número de crédito.
    - df_movimientos: DataFrame que contiene los movimientos relacionados al número de crédito.
    """
    # Configurar el servicio del controlador de Chrome
    s = Service(r"C:\Drivers\chromedriver.exe")  # O "C:/Drivers/chromedriver.exe"

    # Inicializar el controlador de Chrome
    driver = webdriver.Chrome()

    # Abrir la página de inicio de sesión
    driver.get("https://valora.credisoft4.com/Acceso/Login")
    driver.maximize_window()

    # Esperar a que la página de inicio de sesión se cargue completamente
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[contains(@id,'Usuario')]"))
    )

    # Rellenar el usuario y contraseña, luego presionar ENTER
    usuario = driver.find_element(By.XPATH, "//input[contains(@id,'Usuario')]")
    usuario.send_keys("MARCO.OVIEDO" + Keys.TAB+ "Valora24" + Keys.ENTER)

    # Esperar a que la página cargue y el enlace sea clickeable
    WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@href='/Reporte']"))
    )

    # Clickear en el enlace de reportes
    reportes = driver.find_element(By.XPATH, "//span[@class='menu-text'][contains(.,'Créditos')]")
    reportes.click()
    time.sleep(3)

    # Clickear en el botón de listado
    listado = driver.find_element(By.XPATH, "//a[@href='/Credito']")
    listado.click()

    # Rellenar la búsqueda avanzada
    busqueda = driver.find_element(By.XPATH, "//input[@type='search']")
    busqueda.send_keys(numero_credito + Keys.ENTER)

    # Esperar a que desaparezca el elemento "Procesando..."
    WebDriverWait(driver, 100).until(
        EC.invisibility_of_element_located((By.ID, "TablaCreditos_processing"))
    )

    # Le damos click al boton de detalles
    Detalles = driver.find_element(By.XPATH, "//i[contains(@class,'fa fa-bars fa-fw')]") 
    Detalles.click()

    # Clickear en la pestaña de pagos
    Pagos = driver.find_element(By.XPATH, "//a[@data-toggle='tab'][contains(.,'Pagos')]")
    Pagos.click()
    # Esperar a que la tabla se cargue completamente
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//table"))
    )

    # Extraer el código HTML de la página
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')

    # Buscar los encabezados de la tabla
    table_headers = soup.find_all('th', class_='sorting_disabled')

    # Extraer el texto de cada encabezado
    headers = [th.get_text().strip() for th in table_headers]

    # Encontrar la tabla por su identificador único
    tabla = soup.find('table', id='example')

    # Inicializar una lista para almacenar los datos de la tabla
    datos_tabla = []

    # Iterar sobre cada fila de la tabla, excluyendo la primera fila que contiene los encabezados
    for fila in tabla.find_all('tr')[1:]:
        datos_fila = []  # Inicializar una lista para los datos de la fila actual
        # Encontrar todas las celdas en la fila
        celdas = fila.find_all('td')
        # Extraer el texto de cada celda y agregarlo a la lista de datos de la fila
        datos_fila.extend(celda.get_text().strip() for celda in celdas)
        # Agregar la fila de datos a la lista principal
        datos_tabla.append(datos_fila)

    # Convertir la lista de listas en un DataFrame y asignar los encabezados apropiados
    df_pagos = pd.DataFrame(datos_tabla, columns=headers)

        # Extraer las columnas [1,2,3,4,5,19] y excluir la primera fila
    df_pagos = df_pagos.iloc[1:, [1, 2, 3, 4, 5, 19]].copy()

    # Corregir los encabezados
    df_pagos.columns = ['NumeroCredito', 'FechaPago', 'Capital', 'Interes', 'IVA', 'Total']

    # Eliminar las comas (",") de los números
    df_pagos['Capital'] = df_pagos['Capital'].str.replace(',', '')
    df_pagos['Interes'] = df_pagos['Interes'].str.replace(',', '')
    df_pagos['IVA'] = df_pagos['IVA'].str.replace(',', '')
    df_pagos['Total'] = df_pagos['Total'].str.replace(',', '')

    # Eliminar los símbolos "$"
    df_pagos['Capital'] = df_pagos['Capital'].str.replace('$', '').astype(float)
    df_pagos['Interes'] = df_pagos['Interes'].str.replace('$', '').astype(float)
    df_pagos['IVA'] = df_pagos['IVA'].str.replace('$', '').astype(float)
    df_pagos['Total'] = df_pagos['Total'].str.replace('$', '').astype(float)

    # Cambiar el tipo de dato
    df_pagos['NumeroCredito'] = df_pagos['NumeroCredito'].astype(int)
    df_pagos['FechaPago'] = pd.to_datetime(df_pagos['FechaPago'], dayfirst=True)

    # Ordenar la tabla por 'FechaPago' y resetear el índice
    df_pagos = df_pagos.sort_values(by='FechaPago').reset_index(drop=False)
    df_pagos.rename(columns={'index': 'Indice'}, inplace=True)

    # Clickear en la pestaña de movimientos
    Movimientos = driver.find_element(By.XPATH, "//a[@data-toggle='tab'][contains(text(),'Movimientos')]")
    Movimientos.click()

    # Esperar a que la tabla se cargue completamente
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//table"))
    )

    # Extraer el código HTML de la página
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')

    # Encontrar la tabla por su clase
    tabla = soup.find('table', class_='table table-striped table-condensed')

    # Extraer los encabezados de la tabla
    headers = [th.get_text().strip() for th in tabla.find('thead').find_all('th')]
    headers = headers[:-1]  # Excluir el último encabezado que contiene el botón de eliminar

    # Inicializar una lista para almacenar los datos de la tabla
    datos_tabla = []

    # Iterar sobre cada fila de la tabla en el cuerpo de la tabla
    for fila in tabla.find('tbody').find_all('tr'):
        datos_fila = [td.get_text().strip() for td in fila.find_all('td')]
        # El último elemento 'td' contiene un enlace para eliminar, puedes decidir si incluirlo o no
        datos_tabla.append(datos_fila[:-1])  # Excluir el último 'td' que contiene el botón de eliminar

    # Convertir la lista de listas en un DataFrame y asignar los encabezados apropiados
    df_movimientos = pd.DataFrame(datos_tabla, columns=headers)

    # Tomar en cuenta solo las columnas [1,4]
    df_movimientos = df_movimientos.iloc[:, [1, 4]].copy()

    # Eliminar el prefijo "$"
    df_movimientos['Importe'] = df_movimientos['Importe'].str.replace(',', '')
    df_movimientos['Importe'] = df_movimientos['Importe'].str.replace('$', '').astype(float)

    # Cambiar el tipo de dato de 'Fecha Aplicación' a fecha
    df_movimientos['Fecha Aplicación'] = pd.to_datetime(df_movimientos['Fecha Aplicación'], dayfirst=True)

    # Ordenar la tabla por 'Fecha Aplicación' y resetear el índice
    df_movimientos = df_movimientos.sort_values(by='Fecha Aplicación').reset_index(drop=False)
    df_movimientos.rename(columns={'index': 'Indice'}, inplace=True)
    df_movimientos['Indice'] += 1  # Hacer que el índice comience en 1

    # Cerrar el navegador
    driver.quit()

    # Devolver los DataFrames

    return df_pagos, df_movimientos

# Creamos una funcion para generar una tabla de amortizacion
def tabla_amortizacion(informacion):
    # Extraer la información del diccionario
    monto_credito = informacion['MontoCredito']
    plazo = informacion['Plazo']
    pago = informacion['Pago']
    tasa_ordinaria = informacion['TasaOrdinaria']

    # Ahora calculamos el pago amortizable
    pago_amortizable = abs(npf.pmt((tasa_ordinaria*(1+0.16))/24, plazo, monto_credito, 0, 0))

    # Si el pago amortizable es igual al pago entonces (tomamos en cuenta el valor redondeado de los dos pagos, es decir no tomamos en cuenta los decimales)
    if round(pago_amortizable) == round(pago):
        seguro = 0
    else:
        seguro = 10

    # Crear la lista de números de pago y seguro
    numero_pagos = list(range(1, plazo + 1))
    seguro = [seguro] * plazo  # Asumiendo que 'seguro' es un valor predefinido

    # Inicializar las listas para las demás columnas
    saldo = [monto_credito]  # Saldo inicial
    capital = []
    intereses = []
    iva = []
    descuento = []  # Lista vacía para descuentos

    # Aquí actualizamos la lógica para calcular el capital y el descuento
    for i in range(plazo):
        # Calcular los intereses
        interes = saldo[i] * tasa_ordinaria * 15 / 360
        intereses.append(interes)

        # Calcular el IVA sobre los intereses
        iva_interes = interes * 0.16
        iva.append(iva_interes)

        # El total a pagar en esta quincena incluyendo el seguro
        total_pago_quincenal = interes + iva_interes + seguro[i]

        # Calcular el capital como el pago menos el total de intereses, IVA y seguro
        cap = pago - total_pago_quincenal
        capital.append(cap)

        # Calcular el descuento para este periodo
        descuento_actual = pago_amortizable - total_pago_quincenal
        descuento.append(descuento_actual)

        # Actualizar el saldo para el siguiente periodo
        nuevo_saldo = saldo[i] - cap
        if i < plazo - 1:
            saldo.append(nuevo_saldo)

    # Crear el DataFrame con todas las columnas incluyendo Seguro y Descuento
    df_amortizacion = pd.DataFrame({
        'Quincenas': numero_pagos,
        'Saldo': saldo,
        'Capital': capital,
        'Interes': intereses,
        'IVA': iva,
        'Seguro': seguro,
        'Pago': [pago] * plazo
     })
    

    # Devolver el DataFrame
    return df_amortizacion

# Creamos una funcion para detectar el ultimo indice con ceros en movimientos
def encontrar_ultimo_indice_menor_que_pago(df_pagos, pago):
    """
    En esta versión de la función, df_pagos['Total'] < pago crea una serie booleana donde cada elemento es True si el valor de 'Total' en esa fila es menor que pago, y False de lo contrario.
    Luego, si hay al menos una fila donde 'Total' es menor que pago, se encuentra el índice máximo de esas filas para obtener ultimo_indice_menor_que_pago.
    Si no hay tales filas, ultimo_indice_menor_que_pago se establece en None.
    
    Args:
        df_pagos (DataFrame): El DataFrame que contiene los pagos.
        pago (float): El valor de pago a comparar con los valores de 'Total' en el DataFrame.
    
    Returns:
        int or None: El índice máximo de las filas donde 'Total' es menor que pago, o None si no hay tales filas.
    """
    condicion = df_pagos['Total'] < pago
    if condicion.any():  # Verificar si hay al menos una fila donde 'Total' es menor que pago
        ultimo_indice_menor_que_pago = df_pagos[condicion].index.max() + 1
    else:
        ultimo_indice_menor_que_pago = None
    return ultimo_indice_menor_que_pago

# Creamos una funcion para comparar las fehcas del pago y movimientos dada el indice encontrado en la funcion encontrar_ultimo_indice_con_ceros
def comparar_fechas(df_pagos, df_movimientos, ultimo_indice_menor_que_pago):
    # Comparamos las fechas de pago y movimientos
    if ultimo_indice_menor_que_pago is not None:
        fecha_pago = df_pagos['FechaPago'].iloc[int(ultimo_indice_menor_que_pago)-1]
    else:
        fecha_pago = None

    # En movimientos buscamos pero la fecha del ultimo movimiento
    if not df_movimientos['Fecha Aplicación'].empty:
        fecha_movimiento = df_movimientos['Fecha Aplicación'].iloc[-1]
    else:
        fecha_movimiento = None

    # Comparamos las fechas
    if fecha_movimiento == fecha_pago:
        validacion = True
    else:
        validacion = False
    return validacion, fecha_movimiento

# Creamos una funcion para obtener el numero de quincenas dada una fecha
def obtener_numero_quincenas(fecha):
    if fecha is None:
        return None

    fecha_referencia = pd.Timestamp('2021-12-15')
    diferencia = fecha - fecha_referencia

    # Calculamos el número de quincenas
    numero_quincenas = diferencia.days // 15

    # Devolvemos el número de quincenas
    return numero_quincenas

# Creamos una funcion para obtener la informacion de la base de calendario ("C:\Users\alber\Downloads\db_calendario.xlsx")
def informacion_calendario(df_calendario, informacion):
    # Extraemos la institucion del diccionario
    institucion = informacion['Institucion']

    # Cargamos la informacion en un diccionario por la institucion ("Institucion"	"Periodicidad"	"Corte"	"Fecha de corte"	"Fecha de Vencimiento"	"Fecha modificacion")
    df_calendario = df_calendario[df_calendario['Institucion'] == institucion]

    # Devolvemos la base de datos
    return df_calendario

# Creamos una funcion para generar una tabla cruzada
def generar_tabla_cruzada(df_amortizacion, df_pagos, df_calendario, nombre_cliente, numero_credito):
    # De la tabla df_amortizacion vamos a cruzar con la tabla df_pagos para obtener la fecha de pago
    df_cruzado = pd.merge(df_amortizacion, df_pagos[['Indice', 'FechaPago']], left_on='Quincenas', right_on='Indice', how='left')
    # nos quedamos solamente con las columnas Quincenas, Saldo y FechaPago
    df_cruzado = df_cruzado[['Quincenas', 'Saldo', 'FechaPago', 'Pago']]
    # Ahora vamos a cruzar la tabla df_cruzado con la tabla df_calendario para obtener la fecha de corte
    fecha_de_corte = df_calendario['Fecha de corte'].iloc[0]
    # Filtramos la tabla df_cruzado para que solo nos muestre las fechas de pago que sean menores o iguales a la fecha de corte
    df_cruzado = df_cruzado[df_cruzado['FechaPago'] <= fecha_de_corte]
    # Ahora filtramos la fecha más reciente de la columna 'FechaPago'
    # Mantener df_cruzado como DataFrame incluso si solo tiene una fila
    df_cruzado = df_cruzado.loc[[df_cruzado['FechaPago'].idxmax()]]
    # Ahora agregamos columnas nuevas a la tabla df_cruzado que seran todas las columnas de df_calendario
    calendario_dict = df_calendario.iloc[0].to_dict()

    # Itera sobre el diccionario y añade cada par clave-valor como una nueva columna en df_cruzado
    for key, value in calendario_dict.items():
        # Convierte a fecha si es necesario
        if isinstance(value, pd.Timestamp):
            df_cruzado[key] = value.strftime('%d/%m/%Y')
        else:
            df_cruzado[key] = value

    # Cambiamos el formato de las fechas para quitar los minutos y segundos, si es necesario
    if isinstance(df_cruzado['FechaPago'].iloc[0], pd.Timestamp):
        df_cruzado['FechaPago'] = df_cruzado['FechaPago'].apply(lambda x: x.strftime('%d/%m/%Y'))

    # Añadimos la columna de 'Nombre' y 'NumeroCredito'
    df_cruzado['Nombre'] = nombre_cliente.title()
    df_cruzado['NumeroCredito'] = numero_credito

    # Devolvemos la tabla cruzada
    return df_cruzado

# Creamos una funcion para rellenar la carta
def rellenado_carta(doc, fecha_hoy, df_cruzada, idSolicitud,nombre_cliente,fecha,plazo,Numero):
    #print("llego")
    for paragraph in doc.paragraphs:
        if '{{fecha_actual_completa}}' in paragraph.text:
            meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
            fecha_formateada = fecha_hoy.strftime('%d de {} del %Y').format(meses[fecha_hoy.month - 1])
            paragraph.text = paragraph.text.replace('{{fecha_actual_completa}}', fecha_formateada)
        if '{{nombre_cliente}}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{{nombre_cliente}}', nombre_cliente.title())
        if '{{numero_credito}}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{{numero_credito}}', str(df_cruzada['NumeroCredito'].iloc[0]))
        if '{{idSolicitud}}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{{idSolicitud}}', str(idSolicitud))
        if '{{monto_total}}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{{monto_total}}', "{:,.2f}".format(df_cruzada['Saldo'].iloc[0]))
        if '{{descuento}}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{{descuento}}', str(df_cruzada['Pago'].iloc[0]))
        if '{{fecha de vencimiento}}' in paragraph.text:
            meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
            fecha_vencimiento = datetime.strptime(df_cruzada['Fecha de Vencimiento'].iloc[0], '%d/%m/%Y')
            fecha_formateada = fecha_vencimiento.strftime('%d de {} del %Y').format(meses[fecha_vencimiento.month - 1])
            paragraph.text = paragraph.text.replace('{{fecha de vencimiento}}', fecha)
            #print("vea")
        if '{{periodo}}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{{periodo}}', plazo +" "+ Numero +" " + " del 2024")  #str(df_cruzado['Corte'].iloc[0]) , df_cruzado['Periodicidad'].iloc[0]
            #print("vea2")
    #print("hasta aqui")
    archivo_doc = os.path.join(directorio_actual,'plantilla.docx')
    doc.save(archivo_doc)
    #print("pase de aqui")

# Creamos una funcion para generar la carta que creamos en PDF
def generar_pdf(numero_credito):
    archivo_doc = os.path.join(directorio_actual,'plantilla.docx')
    # Creamos un objeto de Word
    word = comtypes.client.CreateObject('Word.Application')
    # Abrimos el documento
    doc = word.Documents.Open(archivo_doc)
    # Guardamos el documento en PDF
    carta_reestructura = os.path.join(directorio_actual, f"cartaReestructura_{numero_credito}")
    doc.SaveAs(carta_reestructura, FileFormat=17)
    #doc.SaveAs(r'C:\Users\alber\Downloads\cartaReestructura_{}.pdf'.format(numero_credito), FileFormat=17)
    # Cerramos el documento
    doc.Close()
    # Cerramos Word
    word.Quit()
