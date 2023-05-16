import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')
import time

print("Este programa normaliza precios de una serie de valores.\n"
      "Lleva todos los valores a los precios de una fecha específica para poder compararlos.\n"
      "Les quita el componente inflacionario utilizando índices de precios.\n"
      "--------------------------------------------------------------------\n"
      "Para funcionar, es necesario que este programa esté en la misma carpeta que el archivo 'input.xlsx'.\n"
      "Una hoja debe llamarse 'indices' y la otra 'serie'.\n"
      "Dentro de estas, deben estar los índices de precios y en la otra, la serie que se quiere normalizar.\n"
      "Los datos deben estar distribuidos como si fueran una tabla")

try:
    indices = pd.read_excel("input.xlsx", sheet_name='indices')
    #serie = pd.read_excel("test.xlsx",sheet_name='Hoja4')
    serie = pd.read_excel("input.xlsx",sheet_name='serie')
    print("Datos obtenidos con éxito")
except:
    print("Parece que el archivo 'input' no esta dentro de la misma carpeta que este programa o los nombres de las hojas no coinciden")

indices.columns = ['fecha','indices']
serie = serie.rename(columns={serie.columns[0]: 'fecha'})

for i in indices,serie:
    i['fecha'] = pd.to_datetime(i['fecha'], format='%Y/%m/%d %H:%M:%S')
    i['fecha'] = i['fecha'].dt.strftime('%m/%y')
df = pd.merge(serie, indices, on='fecha', how='right')

ftarget = input("llevar la serie a precios de: (mm/aa): ")
#ftarget = '01/17'
print("Calculando...")
target = df.loc[df['fecha'] == ftarget, 'indices'].values[0]
df['deflactores'] = target / df.indices

# Multiply all columns in df (except for the 'fecha' column) by 'deflactores'
for col in df.columns:
    if col != 'fecha' and col != 'deflactores':
        df[col] = df[col] * df['deflactores']

# Select only the 'fecha' and multiplied columns in the final output
df = df[['fecha'] + [col for col in df.columns if col != 'fecha' and col != 'deflactores' and col != 'indices']]
df['fecha'] = pd.to_datetime(df['fecha'], format='%m/%y').dt.strftime('%d/%m/%Y')


print("Enviando a excel")
writer = pd.ExcelWriter( "Serie a precios del " + ftarget.replace('/', '-') + ".xlsx" , engine='xlsxwriter') 
df.to_excel(writer, sheet_name='real', index=False)
writer.save()
print("El archivo ha sido guardado en esta misma carpeta")
print("valores convertidos a precios del " + ftarget.replace('/', '-'))
time.sleep(5)