
import pandas as pd
import re
import datetime

def leerexcel(url):
    datwords=pd.read_excel(url)
    palabras = datwords.to_dict()
    return palabras 

def leerENTRADA(url):
    datwords=pd.read_excel(url, sheet_name="Hoja2", usecols="B,C", )  #B y E son las columnas del excel con el número de caso y la subtipología
    print (datwords)
    palabras = datwords.to_dict()
    return palabras    

def getcategorias(url):
    datwords=pd.read_excel(url,1)
    palabras = datwords.to_dict('list')
    return list(palabras.keys())

###MAIN###
url_categorias = "C:/Users/51232209p/Documents/Libro12.xlsx"
url_referencias = "C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/CleanOutput/Libro1.xlsx"
url_resultados = "C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/Correos Procesados/26-04-2022 11-55 Categorizados.xlsx"
categorias = getcategorias(url_categorias)
referencias = leerENTRADA(url_referencias)
resultados = leerexcel(url_resultados)
tabla = {}
for categoria in categorias:
    if categoria not in tabla:
        tabla[categoria] = {}
    for cat in categorias:
        if cat not in tabla[categoria]:
            tabla[categoria][cat] = 0 

print(resultados.keys())
print(referencias.keys())
for id_resultado in range(0,len(resultados["Categoria"])): 
    #print(resultados["Caso"][id_resultado], referencias["Nº Caso"][id_referencia])
    cat_resultado = resultados["Categoria"][id_resultado]
    cat_referencia = resultados["Esperado"][id_referencia]
    tabla[cat_referencia][cat_resultado] = tabla[cat_referencia][cat_resultado] + 1
df_tabla = pd.DataFrame(data= tabla)
x = datetime.datetime.now()
df_tabla.to_excel(str(x.strftime('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/Matrices/%d-%m-%Y %H-%M'))+' Matriz_de_confusion.xlsx')
df_tabla.to_excel('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/Matrices/%d-%m-%Y %H-%M Matriz_Confusion.xlsx')

            
            

