import openpyxl
from openpyxl.drawing.image import Image
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

if "__main__" == __name__:

    comprar = pd.read_csv("compra_semanal.csv", sep =',')
    orders = pd.read_csv("order_details.csv", sep =',')

    libro = openpyxl.Workbook()
    libro.remove(libro.active)

    hoja1 = libro.create_sheet("Reporte ejecutivo")

    df1 = pd.read_csv('compra_semanal.csv')
    df2 = pd.read_csv('order_details.csv')
    df3 = pd.read_csv('pizzas.csv')

    # Creamos gráficos
    ## Ingredientes menos usados
    pequeños = df1.nsmallest(5, 'Porciones')
    plt.figure(figsize=(10, 5))
    ax = sns.barplot(x='Ingredientes', y='Porciones', data=pequeños, palette='rocket_r')
    plt.savefig('Ingredientes_menos_usados.jpg')

    ## Ingredientas más usados
    mayores = df1.nlargest(5, 'Porciones')
    plt.figure(figsize=(10, 5))
    ax = sns.barplot(x='Ingredientes', y='Porciones', data=mayores, palette='rocket_r')
    plt.savefig('Ingredientes_mas_usados.jpg')

    ## Pizzas más pedidas
    plt.figure(figsize=(10, 5))
    ax = sns.countplot(data=df2, x='pizza_id', order=pd.value_counts(df2['pizza_id']).iloc[:5].index)
    ax.set(xlabel='pizza_id', ylabel='Cantidad')
    plt.savefig('Pizzas_mas_pedidas.jpg')

    ## Pizzas más caras
    data = df3.sort_values('price', ascending=False).head(6)
    plt.figure(figsize=(10, 5))
    ax = sns.barplot(x='pizza_id', y='price', data=data, palette='rocket_r')
    ax.set(xlabel='pizza_id', ylabel='price')
    plt.savefig('Pizzas_mas_caras.jpg')

    ## Pizzas más baratas
    data = df3.sort_values('price', ascending=False).tail(6)
    plt.figure(figsize=(10, 5))
    ax = sns.barplot(x='pizza_id', y='price', data=data, palette='rocket_r')
    ax.set(xlabel='pizza_id', ylabel='price')
    plt.savefig('Pizzas_mas_baratas.jpg')

    imagen1 = Image('Ingredientes_menos_usados.jpg')
    imagen2 = Image('Ingredientes_mas_usados.jpg')
    imagen3 = Image('Pizzas_mas_baratas.jpg')
    imagen4 = Image('Pizzas_mas_caras.jpg')
    imagen5 = Image('Pizzas_mas_pedidas.jpg')

    hoja1['A1'] = 'En la siguiente gráfica se muestran los 5 ingredientes menos usados'
    hoja1.add_image(imagen1, 'B2')
    hoja1['T1'] = 'En la siguiente gráfica se muestran los 5 ingredientes más usados'
    hoja1.add_image(imagen2, 'U2')
    hoja1['A31'] = 'En la siguiente gráfica se muestran las 5 pizzas más baratas'
    hoja1.add_image(imagen3, 'B32')
    hoja1['T31'] = 'En la siguiente gráfica se muestran las 5 pizzas más caras'
    hoja1.add_image(imagen4, 'U32')
    hoja1['A61'] = 'En la siguiente gráfica se muestran las 5 pizzas más vendidas'
    hoja1.add_image(imagen5, 'B62')

    libro.save('reporte.xlsx')

    with pd.ExcelWriter("reporte.xlsx", mode="a", engine="openpyxl") as writer:
        comprar.to_excel(writer, sheet_name='Reporte ingredientes', index=False)
        orders.to_excel(writer, sheet_name='Reporte pedidos', index=False)
