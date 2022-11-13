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
    pequeños = df1.nsmallest(20, 'Porciones')
    mayores = df1.nlargest(5, 'Porciones')

    plt.figure(figsize=(10, 5))
    plt.title("Ingredientes menos usados")
    ax = sns.barplot(x='Ingredientes', y='Porciones', data=pequeños, palette='rocket_r')
    plt.xticks(rotation=70)
    plt.savefig('Ingredientes_menos_usados.jpg')

    plt.figure(figsize=(10, 5))
    plt.title("Ingredientes más usados")
    ax = sns.barplot(x='Ingredientes', y='Porciones', data=mayores, palette='rocket_r')
    plt.xticks(rotation=70)
    plt.savefig('Ingredientes_mas_usados.jpg')

    imagen1 = Image('Ingredientes_menos_usados.jpg')
    imagen2 = Image('Ingredientes_mas_usados.jpg')

    hoja1['A1'] = 'En la siguiente gráfica se muestran los 20 ingredientes menos usados'
    hoja1.add_image(imagen1, 'B2')
    hoja1['T1'] = 'En la siguiente gráfica se muestran los 5 ingredientes más usados'
    hoja1.add_image(imagen2, 'U2')

    libro.save('reporte.xlsx')

    with pd.ExcelWriter("reporte.xlsx", mode="a", engine="openpyxl") as writer:
        comprar.to_excel(writer, sheet_name='Reporte ingredientes', index=False)
        orders.to_excel(writer, sheet_name='Reporte pedidos', index=False)