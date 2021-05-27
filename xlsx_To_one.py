import pandas as pd
import glob

Lista_Archivos = []

for i in glob.glob('D:/ReporteActividades/*.xlsx'):
    df = pd.read_excel(i, header=7)
    df = df.iloc[:,[1,3]]
    Lista_Archivos.append(df)

df = pd.concat(Lista_Archivos, ignore_index=True)

guardar = pd.ExcelWriter('Final.xlsx')
df.to_excel(guardar,'Folios')
guardar.save()
