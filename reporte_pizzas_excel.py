# Importo las librerias necesarias para ejecutar el programa
import xlsxwriter
import pandas as pd
import numpy as np
import re, math
from datetime import datetime
import matplotlib.pyplot as plt

def reporte_ejecutivo(excel, ingresos_meses, ingresos_semanas): # Defino la función que crea el reporte ejecutivo
    tipo_letra = 'Times New Roman' # Defino el tipo de letra que voy a utilizar
    hoja1 = excel.add_worksheet('reporte_ejecutivo_2016') # Creo la primera hoja
    hoja1.set_column('B:C',15) # Ensancho las columnas para que se vea bien
    # Defino la lista de los meses
    meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    # Escribo el cabecera de la primera y la segunda columna, es decir, los meses y los ingresos
    hoja1.write(2,1,'MESES', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E'}))
    hoja1.write(2,2,'INGRESOS', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E' })) 
    # Escribo linea a linea la tabla
    for i in range(len(ingresos_meses)):
        hoja1.write(3+i, 1, meses[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
        hoja1.write(3+i, 2, ingresos_meses[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
    # Creo un gráfico de barras verticales con excel 
    ingresos_grafico = excel.add_chart({'type': 'column'})
    # Indico que datos voy a coger para hacer el gráfico
    ingresos_grafico.add_series({'categories': '=reporte_ejecutivo_2016!$B$4:$B$15', 'values': '=reporte_ejecutivo_2016!$C$4:$C$15'})
    ingresos_grafico.set_size({'y_scale': 0.9, 'x_scale': 2}) # Defino el tamaño del gráfico
    # Defino los títulos del gráfico, el eje x y el eje y 
    ingresos_grafico.set_title({'name': 'Ingresos de 2016 de la pizzería Maven según los meses', 'name_font': {'name': tipo_letra, 'bold': True}})
    ingresos_grafico.set_x_axis({'name': 'Meses', 'name_font': {'name': tipo_letra, 'bold': True}})
    ingresos_grafico.set_y_axis({'name': 'Ingresos', 'name_font': {'name': tipo_letra, 'bold': True}})
    ingresos_grafico.set_legend({'none': True}) # Indico que no hay leyenda
    hoja1.insert_chart('E3', ingresos_grafico) # Inserto la gráfica creada en la hoja 
    
    # Defino la lista de las semanas
    semanas = []
    for num in range(len(ingresos_semanas)):
        semanas.append(f'Semana {num}')
    # escribo al cabecera de las columnas, es decir, semanas e ingresos
    hoja1.write(16,1,'SEMANAS', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E'}))
    hoja1.write(16,2,'INGRESOS', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E' })) 
    # Escribo linea a linea la tabla
    for i in range(len(ingresos_semanas)):
        hoja1.write(17+i, 1, semanas[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
        hoja1.write(17+i, 2, ingresos_semanas[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
    # Creo un gráfico de barras verticales con excel
    ingresos_grafico_sem = excel.add_chart({'type': 'line'})
    # Indico qué datos voy a coger para hacer el gráfico
    ingresos_grafico_sem.add_series({'categories': '=reporte_ejecutivo_2016!$B$18:$B$70', 'values': '=reporte_ejecutivo_2016!$C$18:$C$70'})
    ingresos_grafico_sem.set_size({'y_scale': 1.5 , 'x_scale': 3}) # Defino el tamaño del gráfico
    # Defino los títulos del gráfico, el eje x y el eje y 
    ingresos_grafico_sem.set_title({'name': 'Ingresos de 2016 de la pizzería Maven según las semanas', 'name_font': {'name': tipo_letra, 'bold': True}})
    ingresos_grafico_sem.set_x_axis({'name': 'Semanas', 'name_font': {'name': tipo_letra, 'bold': True}})
    ingresos_grafico_sem.set_y_axis({'name': 'Ingresos', 'name_font': {'name': tipo_letra, 'bold': True}})
    ingresos_grafico_sem.set_legend({'none': True}) # Indico que no hay leyenda
    hoja1.insert_chart('E17', ingresos_grafico_sem) # Inserto la gráfica creada en la hoja
    
def reporte_pedidos(excel, pedidos_info): # Defino la función que crea el reporte de pedidos
    tipo_letra = 'Times New Roman' # Defino el tipo de letra que voy a utilizar
    hoja2 = excel.add_worksheet('reporte_pedidos_2016') # Creo la segunda hoja
    hoja2.set_column('B:G',25) # Ensancho las columnas para que se vea bien
    
    # Ajusto los datos para poder utilizarlos
    # Creo una serie con la cantidad de cada pizza que se pide de media en una semana
    # Creo una tabla con la cantidad de cada pizza que se pide de media en una semana en función de los tamaños
    tipos_pizza = pd.Series(pedidos_info['pizza_type_id'].value_counts().divide(len(pedidos_info['week number'].unique().tolist())).apply(np.ceil)).rename_axis('tipos_pizza')
    tabla_tamaños_pizzas = (pd.crosstab(pedidos_info['pizza_type_id'], pedidos_info['size'])/len(pedidos_info['week number'].unique().tolist())).apply(np.ceil).rename_axis('Tipos pizzas').reset_index()
    tipos_pizza = tipos_pizza.reset_index(name = 'cantidad')
    tabla_tamaños_pizzas = tabla_tamaños_pizzas.reindex(columns = ['Tipos pizzas', 'S', 'M', 'L', 'XL', 'XXL'])
    # Creo el gráfico con la librería matplotlib
    plt.figure(1, figsize=(12, 6)) # permite indicar el nº de la figura y las dimensiones (ancho y alto)
    plt.bar(tipos_pizza['tipos_pizza'], tipos_pizza['cantidad'], color = '#01696E') # Defino los datos y de que color quiero que sea la gráfica
    plt.title('Cantidad de pizzas en función de los distintos tipos') # Establezco el título
    plt.xticks(rotation = 90,  fontsize= 5) # Establezco la orientación y el tamaño del nombre de los datos en el eje horizontal
    plt.xlabel('Tipos pizza') # Establezco el título del eje x
    plt.ylabel('Cantidad') # Establezco el título del eje y
    plt.savefig('Cantidad_de_cada_tipo_pizza.png') # Guardo el gráfico como imagen en el propio directorio
    
    grupos = tabla_tamaños_pizzas.iloc[:,0]
    datos_matriz = tabla_tamaños_pizzas.iloc[:,1:].to_numpy().transpose()
    # Establezco los colores de las barras en función del tamaño
    colores = ['#014D50', '#01696E', '#01969D', '#02AEB6', '#02D7E2']
    # Al ser un gráfico de barras apilado no creo la gráfica con la función si no que lo hago directamente aquí
    plt.figure(1, figsize=(12, 6))
    for i in range(datos_matriz.shape[0]):
        # Por cada columna imprimo una barra superpuesta sobre la anterior acerca de cada tipo de pizza y así poder ver el total de pizzas de cada tipo en función del tamaño
        plt.bar(grupos, datos_matriz[i], bottom = np.sum(datos_matriz[:i], axis = 0), color = colores[i], label = tabla_tamaños_pizzas.columns[i+1] )
    # Establezco las características del gráfico
    plt.title('Cantidad de pizzas en función de los distintos tipos y tamaños')
    plt.xticks(rotation = 90,  fontsize= 5)
    plt.ylim(0, 45)
    plt.xlabel('Tipos de pizzas')
    plt.ylabel('Cantidad')
    plt.legend() # Creo una leyenda para saber que tamaño representa cada color
    plt.savefig('Tipos_pizzas_tamaño.png') # Guardo el gráfico como imagen en el propio directorio
    # Escribo la cabecera de las columnas, es decir, los distintos tipos de pizza y la cantidad de cada una
    hoja2.write(2,3,'TIPOS PIZZAS', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E'}))
    hoja2.write(2,4,'CANTIDAD/SEMANA', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E' })) 
    # Escribo linea a linea la tabla 
    for i in range(len(tipos_pizza)):
        hoja2.write(3+i, 3, tipos_pizza['tipos_pizza'].iloc[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
        hoja2.write(3+i, 4, tipos_pizza['cantidad'].iloc[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
    # Escribo la cabecera de las columnas, es decir, los distintos tipos de pizza y la cantidad de cada una en función del tamaño 
    hoja2.write(36,1,'TIPOS PIZZA', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E'}))
    hoja2.write(36,2,'CANDIDAD S/SEMANA', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E' })) 
    hoja2.write(36,3,'CANDIDAD M/SEMANA', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E'}))
    hoja2.write(36,4,'CANDIDAD L/SEMANA', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E' })) 
    hoja2.write(36,5,'CANDIDAD XL/SEMANA', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E'}))
    hoja2.write(36,6,'CANDIDAD XXL/SEMANA', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E' })) 
    # Escribo linea a linea la tabla
    for i in range(len(tabla_tamaños_pizzas)):
        hoja2.write(37+i, 1, tabla_tamaños_pizzas['Tipos pizzas'].iloc[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
        hoja2.write(37+i, 2, tabla_tamaños_pizzas['S'].iloc[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
        hoja2.write(37+i, 3, tabla_tamaños_pizzas['M'].iloc[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
        hoja2.write(37+i, 4, tabla_tamaños_pizzas['L'].iloc[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
        hoja2.write(37+i, 5, tabla_tamaños_pizzas['XL'].iloc[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
        hoja2.write(37+i, 6, tabla_tamaños_pizzas['XXL'].iloc[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
    # Pego ambos gráficos creados en la hoja de excel 
    hoja2.insert_image('I5', 'Cantidad_de_cada_tipo_pizza.png')
    hoja2.insert_image('I39', 'Tipos_pizzas_tamaño.png')

def reporte_ingredientes(excel, diccionario): # Defino la función que crea el reporte de pedidos
    tipo_letra = 'Times New Roman' # Defino el tipo de letra que voy a utilizar
    hoja3 = excel.add_worksheet('reporte_ingredientes_2016') # Creo la tercera hoja
    hoja3.set_column('B:C',25) # Ensancho las columnas para que se vea bien
    # Defino las listas que continen los datos
    ingredientes = list(diccionario.keys())
    cantidad = list(diccionario.values())
    # Escribo las cabeceras de la tabla, es decir, los ingredientes y sus cantidades
    hoja3.write(2,1,'INGREDIENTES', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E'}))
    hoja3.write(2,2,'CANTIDAD', excel.add_format({'align': 'center', 'font_name': tipo_letra, 'bold': True, 'border': 1, 'border_color': '#01696E' })) 
    # Escribo linea a linea la tabla
    for i in range(len(ingredientes)):
        hoja3.write(3+i, 1, ingredientes[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
        hoja3.write(3+i, 2, cantidad[i], excel.add_format({'align': 'center', 'font_name': tipo_letra, 'border': 1, 'border_color': '#01696E'}))
    # Creo un gráfico de barras verticales con excel
    ingredientes_grafico = excel.add_chart({'type': 'column'})
    # Indico qué datos voy a coger para hacer el gráfico
    ingredientes_grafico.add_series({'categories': '=reporte_ingredientes_2016!$B$4$:$B$68$', 'values': '=reporte_ingredientes_2016!$C$4:$C$68'})
    ingredientes_grafico.set_size({'y_scale': 2, 'x_scale': 4}) # Defino el tamaño del gráfico
    # Defino los títulos del gráfico, el eje x y el eje y 
    ingredientes_grafico.set_title({'name': 'Cantidad de ingredientes necesarios en una semana', 'name_font': {'name': tipo_letra, 'bold': True}})
    ingredientes_grafico.set_x_axis({'name': 'Ingredientes', 'name_font': {'name': tipo_letra, 'bold': True}})
    ingredientes_grafico.set_y_axis({'name': 'Proporción de pizzas tamaño S', 'name_font': {'name': tipo_letra, 'bold': True}})
    ingredientes_grafico.set_legend({'none': True}) # Indico que no hay leyenda
    hoja3.insert_chart('E3', ingredientes_grafico) # Inserto la gráfica creada en la hoja
    return

def extract(): # Creo la función que extrae los datos (la E de mi ETL)
    # Guardo cada archivo de tipo .csv en un dataframe, teniendo en cuenta el encoding y el separador necesario para que lo pueda leer sin problemas
    order_details = pd.read_csv('order_details.csv', sep = ';')
    orders = pd.read_csv('orders.csv', sep = ';')
    pizzas = pd.read_csv('pizzas.csv', sep = ',')
    pizza_types = pd.read_csv('pizza_types.csv', sep = ',', encoding = 'unicode_escape')
    return order_details, orders, pizzas, pizza_types

def arreglar_dataframes(orders, order_details, pizzas, pizza_types):  # Creo la función que va a formatear los datos para devolverlos como queremos 
    # Establezco todos los formatos de fechas para ir transformando la columna de la fecha
    formatos = ['%B %d %Y', '%b %d %Y', '%Y-%m-%d', '%d-%m-%y %H:%M:%S', '%A,%d %B, %Y', '%a %d-%b-%Y']
    # Junto los df de los orders y los orders_details para poder trabajar con todo a la vez
    order_details = order_details.merge(orders, on = 'order_id')
    # Creo un dataframe con los valores en los que el id de la pizza no es nan
    order_details = order_details[order_details['pizza_id'].notna()]
    # Limpio todos los datos mal puestos de la columna del id de cada pizza sustituyendo los guiones por barras bajas
    # los espacios por barras bajas, los @ por a, los 3 por e y los ceros por o
    order_details['pizza_id'] = order_details['pizza_id'].apply(lambda x: re.sub('-', '_', x))
    order_details['pizza_id'] = order_details['pizza_id'].apply(lambda x: re.sub(' ', '_', x))
    order_details['pizza_id'] = order_details['pizza_id'].apply(lambda x: re.sub('@', 'a', x))
    order_details['pizza_id'] = order_details['pizza_id'].apply(lambda x: re.sub('3', 'e', x))
    order_details['pizza_id'] = order_details['pizza_id'].apply(lambda x: re.sub('0', 'o', x))
    # Limpio todos los datos de la columna de quantity y para ello primero relleno todos los nan con un 1 
    # y para los One / one pongo un 1 y para los two / Two pongo un 2, por último cambio los negativos a positivos
    order_details['quantity'] = order_details['quantity'].fillna(1)
    order_details['quantity'] = order_details['quantity'].apply(lambda x: re.sub('one', '1',  str(x).lower() ))
    order_details['quantity'] = order_details['quantity'].apply(lambda x: re.sub('two', '2', x.lower()))
    order_details['quantity'] = order_details['quantity'].apply(lambda x: abs(int(x))) 
    # Ordeno el datafrmae según el id del pedido para poder establecer las fechas que falten 
    order_details_ord = order_details.sort_values('order_details_id', ascending = True).reset_index()
    order_details_ord['week number'] = None  # Creo una columna que va a guardar el número de la semana
    order_details_ord['month number'] = None
    for fecha in range(0,len(order_details_ord['date'])):
        # Para cada pedido formateo el tipo del dato de la fecha
        try:
            # Intento arreglar la fecha con la funcion fromtimestamp de la libreria datetime
            fecha_arreglada = float(order_details_ord['date'].iloc[fecha])
            fecha_final = datetime.fromtimestamp(fecha_arreglada)
            order_details_ord['date'].iloc[fecha] = fecha_final
        # si salta un error es porque o bien la fecha es un nan o no está en el formato indicado 
        except ValueError as error:
            tipo = 0
            # Para cada uno de los formatos que puede tener una fecha
            while 0 <= tipo < len(formatos):
                try: # Intentamos transformar la fecha en ese formato
                    # Si podemos transformarlo, lo transformamos y lo cambiamos en nuestro dataframe
                    fecha_final = datetime.strptime(order_details_ord['date'].iloc[fecha], formatos[tipo])
                    order_details_ord['date'].iloc[fecha] = fecha_final
                    tipo = -1 # Salgo del bucle
                except: 
                    # Si no es ese tipo probamos con el siguiente hasta probar con todos los tipos
                    tipo += 1
            if tipo != -1: # Si el tipo no es ninguno de los definifos en la lista
                if fecha == 0: # La fecha es la primera del datafrmae establezco como la fecha el uno de enero del
                    # año ya que sabemos que es el primer pedido realizado en el año por haberlo ordenado previamente 
                    fecha_final = datetime.strptime('01-01-2016', '%d-%m-%Y')
                else:
                    # Si no es el primer pedido del año, entonces establecemos la misma fecha que el pedido anterior
                    fecha_final = order_details_ord['date'].iloc[fecha-1]
                order_details_ord['date'].iloc[fecha] = fecha_final
        # Para cada pedido guardo en que semana se realizó
        numero_semana = int(order_details_ord['date'].iloc[fecha].strftime('%W'))
        numero_mes = int(order_details_ord['date'].iloc[fecha].strftime('%m'))
        order_details_ord['week number'].iloc[fecha] = numero_semana
        order_details_ord['month number'].iloc[fecha] = numero_mes
    pizzas = pizzas.merge(pizza_types, on = 'pizza_type_id') # Junto los df con los detalles de cada pizza
    pedidos_informacion = order_details_ord.merge(pizzas, on = 'pizza_id') #Junto todos los datos en uno mismo
    return order_details_ord, pizzas, pedidos_informacion # Devuelvo los dfs transformador y con los datos formateados

def ingredientes_pizzas(pizzas): # Creo la función que guarda en un dataframe qué ingredientes tiene cada pizza
    tipos_pizza = pizzas['pizza_id'].unique().tolist() # Establezco los tipos de pizza que existen en el restaurante
    ingredientes = set() # Creo un set en el que voy a añadir los ingredientes que se pueden añadir a las pizzas
    for i in range(len(pizzas)): 
        # Recorro el dataframe añadiendo a mi set los ingredientes de forma única
        for ing in [ing.strip() for ing in pizzas['ingredients'].iloc[i].split(',')]:
            ingredientes.add(ing)
    matriz = [] # Creo la matriz que va a definir mi dataframe de las pizzas con sus ingredientes
    for fila in range(len(pizzas)): # Para cada fila del df de las pizzas
        tipo = [] 
        for ingrediente in ingredientes:
            # Para cada tipo de pizza guardo un 1 si contiene al ingrediente o un 0 si no
            if ingrediente in [ing.strip() for ing in pizzas['ingredients'].iloc[fila].split(",")]:
                tipo.append(1)
            else:
                tipo.append(0)
        matriz.append(tipo)
    matriz = np.array(matriz)
    # Creo el dataframe en el que las filas son los tipos de pizza (distinguiendo por tamaño) y las columnas son cada uno de los ingredientes
    ingre_pizzas = pd.DataFrame(data = matriz, index = tipos_pizza, columns = [ingrediente for ingrediente in ingredientes])
    return ingre_pizzas

def ingredientes_semana(order_details, ingre_pizzas): # Creo la función que me calcula la cantidad de cada ingrediente por semana
    semanas = []
    multiplicadores = {'s': 1, 'm': 2, 'l': 3} 
    # Establezco que si la pizza es mediana necesite el doble de cantidad que una pizza pequeña y que si es grande necesite el triple
    orders_week = order_details.groupby(by= 'week number') # Divido todos los pedidos por semanas
    for week in orders_week: 
        # Para cada semana creo una copia de la tabla que indica que ingredientes tiene cada pizza para completarla con la semana en cuestion 
        df_nuevo = ingre_pizzas.copy()
        # Creo una tabla de contingencia para contar el número de pizzas de cada tipo que hay según el número por pedido
        contingencia = pd.crosstab(week[1].pizza_id, week[1].quantity)
        indices = contingencia.index # Guardo los nombres de las pizzas
        for indice in indices: 
            # Para cada tipo de pizza guardo el multiplicador, es decir, el valor por el que voy a multiplicar en función del tamaño de la pizza 
            multiplicador = multiplicadores[indice[-1]]
            numero = 0
            columnas = contingencia.columns
            # Calculo el número de pizzas de cada tipo por semana en proporción a si todas fuesen de tamaño pequeño 
            for columna in columnas: 
                numero += (contingencia[columna][indice]) * columna
            # Multiplico la fila del dataframe que guarda que ingredientes contiene cada pizza por el número de pizzas que tengo que hacer esa semana
            df_nuevo.loc[indice] = df_nuevo.loc[indice].mul(multiplicador*numero)
        # Devuelvo una lista con todos los df de los ingredientes por semana
        semanas.append(df_nuevo)
    return semanas, orders_week

def crear_recuento_semana(semanas): # Creo la función que me hace un recuento de los ingredientes por semana
    diccionario = {} 
    for semana in semanas: 
        columnas = semana.columns
        # Para cada semana hago el recuento de cada ingrediente y lo añado  a diccionario que contiene como claves los ingredientes y como valores
        # una lista con la cantidad del ingrediente necesaria para cada semana en función del número de pizzas (cada posicion de la lista es una semana)
        for columna in columnas:
            if columna not in diccionario:
                diccionario[columna] = []
            contador = semana[columna].sum()
            diccionario[columna].append(contador)
    # Devuelvo el diccionario con todos los ingredientes y sus respectivas cantidades
    dict_medias = {}
    for clave in diccionario: # Para cada ingrediente predigo la cantidad necesaria en una semana 
        # Para predecir los ingredientes necesarios por semana hago la media de los necesitados en cada semana del año
        dict_medias[clave] = math.ceil(np.array(diccionario[clave]).mean()) 
    return dict_medias

def crear_finanzas(pedidos_info): # Defino la función que calcula los ingresos en función de los meses y las semanas
    orders_month = pedidos_info.groupby(by= 'month number') # Divido todos los pedidos por meses
    ingresos_meses = []
    # Calculo los precios de cada mes
    for mes in orders_month:
        ingresos_meses.append(round(mes[1]['price'].sum(),2))
    orders_week = pedidos_info.groupby(by= 'week number') # Divido todos los pedidos por semanas
    ingresos_semanas = []
    # Calculo los precios de cada semana
    for semana in orders_week:
        ingresos_semanas.append(round(semana[1]['price'].sum(),2))
    return ingresos_meses, ingresos_semanas

def extract(): # Creo la función que extrae los datos (la E de mi ETL)
    # Guardo cada archivo de tipo .csv en un dataframe, teniendo en cuenta el encoding y el separador necesario para que lo pueda leer sin problemas
    order_details = pd.read_csv('order_details.csv', sep = ';')
    orders = pd.read_csv('orders.csv', sep = ';')
    pizzas = pd.read_csv('pizzas.csv', sep = ',')
    pizza_types = pd.read_csv('pizza_types.csv', sep = ',', encoding = 'unicode_escape')
    return order_details, orders, pizzas, pizza_types

def transform(orders, order_details, pizzas, pizza_types):
    # Llamo a la función que me formatea los dataframes y los datos 
    order_details, pizzas, pedidos_info = arreglar_dataframes(orders, order_details, pizzas, pizza_types) 
    ingre_pizzas = ingredientes_pizzas(pizzas) # Llamo a la función que me genera el dataframe con los ingredientes que contiene cada pizza
    # Llamo a la función que me genera un df por cada semana con los ingredientes necesarios para dicha semana
    semanas, orders_week = ingredientes_semana(order_details, ingre_pizzas) 
    # Creo el diccionario que contiene para cada ingrediente una lista con la cantidad necesaria para cada semana
    diccionario = crear_recuento_semana(semanas)
    return diccionario, pedidos_info

def load(diccionario, pedidos_info):
    excel = xlsxwriter.Workbook('reporte_pizzas.xlsx') # Creo el documento como un objeto 'Workbook'
    ingresos_meses, ingresos_semanas = crear_finanzas(pedidos_info) # Llamo a crear_finanzas para tener los ingresos de los meses y las semanas
    # Añado las hojas con los respectivos informes
    reporte_ejecutivo(excel, ingresos_meses, ingresos_semanas)
    reporte_pedidos(excel, pedidos_info)
    reporte_ingredientes(excel, diccionario)
    excel.close() # Cierro el documento para que quede guardado en el directorio
    
if __name__=="__main__":
    order_details, orders, pizzas, pizza_types = extract() # Llamo a extract para extraer todos los datos 
    diccionario, pedidos_info = transform(order_details, orders, pizzas, pizza_types)
    # LLamo a transform para trabajar los datos y poder operar con ellos
    load(diccionario, pedidos_info) # Llamo a load para crear el documento excel con los datos y guardarlo
