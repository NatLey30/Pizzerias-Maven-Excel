import pandas as pd

def archivo(fichero):
    print(f'ANALISIS DE DATOS DEL FICHERO {fichero}')
    df = pd.read_csv(fichero, sep = ',', encoding = 'LATIN1')

    nans = str(df.isna().sum().sum())
    print(f'Nans totales: {nans}')

    nulls = str(df.isnull().sum().sum())
    print(f'Nans totales: {nulls}')

    columnas = df.columns.values
    for colum in range(len(columnas)):
        nombre = columnas[colum]
        tipo_col = str(df[columnas[colum]].dtype)
        print(f'     La columna "{nombre}" es del tipo: {tipo_col}')

if __name__ == "__main__":

    archivo('order_details.csv')
    print('---------------------------------------------------------------')
    archivo('orders.csv')
    print('---------------------------------------------------------------')
    archivo('data_dictionary.csv')
    print('---------------------------------------------------------------')
    archivo('pizzas.csv')
    print('---------------------------------------------------------------')
    archivo('pizza_types.csv')