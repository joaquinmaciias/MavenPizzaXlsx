import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re
import xlsxwriter


def load_Excel(dict_pizza_orders: dict, dict_ingredients_weekly: dict):

    Excel = xlsxwriter.Workbook('MavenPizza.xlsx') # Creamos el fichero excel

    worksheet_DataFrames = Excel.add_worksheet(name = 'DataFrames') # Añadimos una pagina al fichero

    worksheet_DataFrames.set_column('A:A', 20) # Añadimos una columna
    worksheet_DataFrames.set_column('B:B', 20) # Añadimos una columna

    bold = Excel.add_format({'bold': True})

    worksheet_DataFrames.write('A1', 'Pizzas Ordered', bold)
    worksheet_DataFrames.write('B1', 'Quantity', bold)


    row = 3
    col = 0

 # Iterate over the data and write it out row by row.
    for pizza in dict_pizza_orders:

        contenidoA = 'Pizza ' + pizza
        worksheet_DataFrames.write(row, col, contenidoA)

        contenidoB = dict_pizza_orders[pizza]
        worksheet_DataFrames.write(row, col + 1, contenidoB)

        row += 1

        
    worksheet_DataFrames.write(row + 1, col, 'Total Pizzas',bold)
    worksheet_DataFrames.write(row + 1, col + 1, f'=SUM(B3:B{str(row - 1)})')

    worksheet_DataFrames.set_column('D:D', 20) # Añadimos una columna
    worksheet_DataFrames.set_column('E:E', 20) # Añadimos una columna

    bold = Excel.add_format({'bold': True})

    worksheet_DataFrames.write('D1', 'Ingredients Week', bold)
    worksheet_DataFrames.write('E1', 'Quantity', bold)


    row = 3
    col = 3

 # Iterate over the data and write it out row by row.
    for ingredient in dict_ingredients_weekly:

        contenidoA = ingredient
        worksheet_DataFrames.write(row, col, contenidoA)

        contenidoB = int(dict_ingredients_weekly[ingredient])
        worksheet_DataFrames.write(row, col + 1, contenidoB)

        row += 1

        
    worksheet_DataFrames.write(row + 1, col, 'Total Ingredients',bold)
    worksheet_DataFrames.write(row + 1, col + 1, f'=SUM(E3:E{str(row - 1)})')


    worksheet_graph_pizza = Excel.add_worksheet(name = 'Graph Pizza') # Añadimos una pagina al fichero
    worksheet_graph_pizza.write(0, 0, 'Pizzas Graph',bold)

    worksheet_graph_pizza.insert_image('C3', 'Pizzas_pedidas.jpg')

    
    worksheet_graphs_ingredient = Excel.add_worksheet(name = 'Graph Ingredient') # Añadimos una pagina al fichero
    worksheet_graphs_ingredient.write(0, 0, 'Ingredients Graph',bold)

    worksheet_graphs_ingredient.insert_image('C3', 'Ingredientes_semanales.jpg')

    Excel.close()

def details_csv(df_name:str,df: pd.DataFrame):

    # Analisis calidad de los datos

    print(f'CSV name is: {df_name} -->\n')

    print(df.info())
    print('\n')

    print(df.head(10))
    print('\n')

    print('Dimensiones del Csv\n')

    print(df.shape)
    print('\n')

    print('Columnas del Csv\n')

    print(df.columns)
    print('\n')

    print('Datos nulos del Csv\n')

    print(df.isnull().sum())

    print('\n')

    print('----------------------------------------------------------')


def extract(csv_file: str, sep: str) -> pd.DataFrame:  # Extraemos los datos de los csv

    df = pd.read_csv(csv_file, sep = sep)

    return df

def clear_data(df:pd.DataFrame):

    df = df.dropna()

    return df

def transform_orders(df: pd.DataFrame):  # Transformamos los datos de los dataframes

    pizza_rep = dict()
    df_pizza_quantity = df.loc[:, ['pizza_id', 'quantity']]  # Nos guardamos las columnas que nos interesan del dataframe
    list_pizza_quantity = df_pizza_quantity.values.tolist()  # Pasamos el dataframe a una lista

    # Eliminamos con expresiones regulares el tamaño de la pizza que acompaña al nombre de la pizza


    for i in range(len(list_pizza_quantity)):

        list_pizza_quantity[i][1] = re.sub("One","1",list_pizza_quantity[i][1])
        list_pizza_quantity[i][1] = re.sub("one","1",list_pizza_quantity[i][1])
        list_pizza_quantity[i][1] = re.sub("-1","1",list_pizza_quantity[i][1])
        list_pizza_quantity[i][1] = re.sub("two","2",list_pizza_quantity[i][1])
        list_pizza_quantity[i][1] = re.sub("Two","2",list_pizza_quantity[i][1])

        list_pizza_quantity[i][0] = re.sub("[ -]","_", list_pizza_quantity[i][0])
        list_pizza_quantity[i][0] = re.sub("@","a", list_pizza_quantity[i][0])
        list_pizza_quantity[i][0] = re.sub("0","o", list_pizza_quantity[i][0])
        list_pizza_quantity[i][0] = re.sub(" ","_", list_pizza_quantity[i][0])
        list_pizza_quantity[i][0] = re.sub("3","e", list_pizza_quantity[i][0])
        list_pizza_quantity[i][0] = re.sub('_[a-z]$', '', list_pizza_quantity[i][0])


    # Creamos un diccionario con claves los nombres de las pizzas y sus valores el número de veces que fueron pedidas.

    for i in range(len(list_pizza_quantity)):
        try:
            pizza_rep[list_pizza_quantity[i][0]] += 1
        except:
            pizza_rep[list_pizza_quantity[i][0]] = 1

    return pizza_rep


def transform_ingredients(df: pd.DataFrame, dict_orders:dict):

    dict_ingredients = dict()
    df_pizza_ingredients = df.loc[:, ['pizza_type_id','ingredients']]
    list_pizza_ingredients = df_pizza_ingredients.values.tolist()

    for i in range(len(list_pizza_ingredients)):
        list_pizza_ingredients[i][1] = list_pizza_ingredients[i][1].replace(', ', ',')
        list_pizza_ingredients[i][1] = re.findall(r'([^,]+)(?:,|$)', list_pizza_ingredients[i][1])
        # list_pizza_ingredients[i] = [nombre pizza,[ingrediente(1),...,ingrediente(k)]]

    for typepizza in range(len(list_pizza_ingredients)):

        for ingredients in range(len(list_pizza_ingredients[typepizza][1])):

            number_order_pizza = dict_orders[list_pizza_ingredients[typepizza][0]]  # Numero de veces que se pedio la pizza con dicho ingrerdiente

            try:
                dict_ingredients[list_pizza_ingredients[typepizza][1][ingredients]] = number_order_pizza + dict_ingredients[list_pizza_ingredients[typepizza][1][ingredients]]
            except:
                dict_ingredients[list_pizza_ingredients[typepizza][1][ingredients]] = number_order_pizza

    return dict_ingredients


def load_graphic_pizzas(pizza_rep: dict) -> pd.DataFrame :

    df  = pd.DataFrame()

    list_type_pizza = []
    list_quantity = []

    for key in pizza_rep:
        list_type_pizza.append(key)
        list_quantity.append(pizza_rep[key])


    df['pizza_id'] = list_type_pizza
    df['quantity'] = list_quantity

    plt.rcParams.update({'font.size': 24})
    plt.figure(figsize=(50, 25))
    ax = sns.barplot(x='pizza_id', y='quantity', data=df,palette='rocket_r')  # Gráfica pizzas vendidas
    ax.set_xticklabels(ax.get_xticklabels(),rotation=70)
    ax.set_title('Pizzas pedidas')
    plt.savefig('Pizzas_pedidas.jpg') # Guardamos la gráfica como una imagen jpg




def load_graphic_ingredients(dict_ingredients_weekly: dict,):

    # Dict -> DataFrame

    df = pd.DataFrame()
    ingredients = []
    quantity = []

    for key in dict_ingredients_weekly:
        quantity.append(key)
        ingredients.append(dict_ingredients_weekly[key])

    df['Cantidad'] = quantity
    df['Ingredientes'] = ingredients

    plt.rcParams.update({'font.size': 24})
    plt.figure(figsize=(50, 25))
    ax = sns.barplot(x='Cantidad', y='Ingredientes', data=df,palette='rocket_r') # Gráfica ingredientes semanales
    ax.set_xticklabels(ax.get_xticklabels(),rotation=70)
    ax.set_title('Ingredientes semanales')
    plt.savefig('Ingredientes_semanales.jpg') # Guardamos la gráfica como una imagen jpg

if __name__ == '__main__':


    # Seguiremos el preceso ETL

    # Extract

    date_orders = extract('orders.csv',';')
    order_details = extract('order_details.csv',';')
    pizzas = extract('pizzas.csv',',')
    pizza_types = extract('pizza_types.csv',',')

    # Analisis

    details_csv('orders.csv',date_orders)
    details_csv('order_details.csv',order_details)
    details_csv('pizzas.csv',pizzas)
    details_csv('pizza_types.csv',pizza_types)

    # Clear data

    date_orders_clean = clear_data(date_orders)
    order_details_clean = clear_data(order_details)
    pizzas_clean = clear_data(pizzas)
    pizza_types_clean = clear_data(pizza_types)

    
    # Transform
    
    dict_pizza_orders = transform_orders(order_details_clean)
    dict_ingredients_anual = transform_ingredients(pizza_types_clean,dict_pizza_orders)
    
    # Ingredients semanales

    dict_ingredients_weekly = dict()

    for key in dict_ingredients_anual:

        dict_ingredients_weekly[key] = dict_ingredients_anual[key] / 12 * 4

    # Load data on the Excel file

    load_graphic_ingredients(dict_ingredients_weekly)
    load_graphic_pizzas(dict_pizza_orders)

    load_Excel(dict_pizza_orders, dict_ingredients_weekly)
