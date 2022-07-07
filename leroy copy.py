import re
import pandas as pd
from pandas import ExcelWriter
import spacy
import requests
import string
import math
from spacy import tokens
from spacy.matcher import Matcher
from spacy.lang.es import Spanish
from spacy.lang.en import English
from spacy.lang.it import Italian
from spacy.lang.fr import French
#from langdetect import detect
import pathlib
import datetime
import openpyxl 
#from xlrd import open_workbook      # Librería para leer archivos excel .xls como si fuese un libro en bloque
import copy
from os import remove 

class Categoria:
    def __init__(self):
        self.palabras = 0 
        self.tfidfs = 0

categorias = {}     # Variable global para almacenar los datos necesarios del excel
motivo = {}
categorias_finales = []
ids_finales = []
cat_esperada=[]

def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Print New Line on Complete
    if iteration == total: 
        print()



def WriteExcel(id,cat,mot,exp):

    print("Escribimos excel")
    x = datetime.datetime.now()
    do = pd.DataFrame({'ID': id, 'Categoria': cat, 'Motivo': mot, 'Esperado': exp}) 
    writer = ExcelWriter(str(x.strftime('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/Correos Procesados/%d-%m-%Y %H-%M'))+' Categorizados.xlsx')
    do.to_excel(writer, 'Hoja de datos', index=False)
    writer.save()           


def calcular_tfidf(palabra, correo, correos_total):
    numero_palabras = len(correo)       # Variable con el número total de palabras analizadas en ese correo 
    numero_correos= len(correos_total)
    contador_tf = 0
    contador_idf = 0
    # Bucle para contar dentro de ese correo el número de veces que aparece el lema, en la variable contador_tf. 
    for otra_palabra in correo: 
           
        if otra_palabra.lemma_ == palabra:
            contador_tf = contador_tf + 1
    if(numero_palabras>0):
        tf = contador_tf/(numero_palabras) 
        
    else:
        return 0
    
    for otro_correo in correos_total:
        for otra_palabra in otro_correo:
            
            if otra_palabra.lemma_== palabra:
                contador_idf = contador_idf + 1
                break
    if(contador_idf==0):
        contador_idf=1
    idf = numero_correos/ float(contador_idf)
    
    # Fórmula del cálculo total tf-idf 
    tfidf = tf * math.log(idf)
    #valor= contador_tf* tfidf
    return tfidf

def category(correo, total_correos):
    #contador=[0] * 7
    #cont_tfidf=[0]* 7
    valores={}
    valores["Sin Clasificar"]=Categoria()
    url_keywords='C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/categorías.xlsx'
    datwords=pd.read_excel(url_keywords,sheet_name='Hoja2')
    for token in correo:
         if (not token.is_stop 
            and not token.is_punct 
            and not token.text == ' ' 
            and not token.like_url 
            and not token.pos_=="PROPN"
            and token.lemma_.lower()
            and token.lemma_ != " "):
            tfidf=0
            tfidf=calcular_tfidf(token.lemma_,correo,total_correos)
            for i in datwords.index:
                try:
                    listwords=datwords.iloc[:,i].to_list()
                except:
                    break
                if token.lemma_.lower() in listwords or token.text.lower() in listwords:
                    categoria=datwords.columns.values[i]
                    if not (categoria in valores):
                        valores[categoria]=Categoria()
                    if categoria=="Stock":
                        valores[categoria].tfidfs+=tfidf*2
                    else:
                        valores[categoria].tfidfs+=tfidf

                    valores[categoria].palabras+=1
                else:
                    valores["Sin Clasificar"].tfidfs=tfidf*0.0001    
                    valores["Sin Clasificar"].palabras+=1


    
    #print(contador)
    #print(cont_tfidf)
    for val in valores:
            valores[val].tfidfs= valores[val].tfidfs* valores[val].palabras
    max_val=valores["Sin Clasificar"].tfidfs
    pos=0
    index=0
    for val in valores:
        if (valores[val].tfidfs > max_val):
            max_val= valores[val].tfidfs
            pos=index
        index+=1
    key_list = list(valores.keys())
 
    return key_list[pos]

def ReadExcel(url_excel):       

    
    emails_total = []
    ids = []
    info = {}
    expectedCategory=[]

    nlp = spacy.load('es_core_news_sm')

    df = pd.read_excel(url_excel, 0)
    #Eliminamos del dataframa la fila dónde el campo descripción está vacío
    #df = df.dropna(subset=["Descripción (Objeto) (Correo electrónico)"])

    for i in range(1000):
        
        id = df.iloc[i,0]               #Columna 1 -> 'ID'
        correos = df.iloc[i,2]          #Columna 2 -> 'Descripción'
        expectedCategory.append(df.iloc[i,1])
        
        #mensaje_lower = "".join(correos.lower()).split("escribió") #CORREOS
        #mensajecss = re.sub('(?s)<style>(.*?)</style>',' ',mensaje_lower[0]) #CORREOS
        mensaje_lower = "".join(str(correos).lower()) #CORREOS
        mensajecss = re.sub('(?s)<style>(.*?)</style>',' ',mensaje_lower) #CORREOS

        #mensaje_lower = "".join(correostring.lower()).split("Comentarios:") #FORMULARIOS
        #mensajecss = re.sub('(?s)<style>(.*?)</style>',' ',mensaje_lower[1]) #FORMULARIOS
        mensajehtml = re.sub('<.*?>|&([a-z0-9]+|#[0-9]{1,6}|#x[0-9a-f]{1,6});',' ',mensajecss)
        em = mensajehtml.replace(r'<[^>]+>', ' ').strip()      #quitamos etiquetas html
        men = re.sub(r'(https|http)?:\/\/(\w|\.|\/|\?|\=|\&|\%)*\b', ' ', em, flags=re.MULTILINE)

        men = re.sub("\d+", ' ', men)   
        men = re.sub(r"(?<=[a-z])\r?\n"," ", men)
        men = re.sub("\."," ", men)
        men = men.translate(str.maketrans(' ',' ',string.punctuation))    
        men = men.strip()
        
        doc = nlp(men)

        ids.append(id)
        emails_total.append(doc)
        info = dict(zip(ids,emails_total))      #diccionario con ids de los correos y los correos

    print("CATEGORIZAMOS")
    contador_cat=[0]*8
    contador =0
    aciertos=0
    for i,c in info.items():
        if contador <999: 
            printProgressBar(contador,1000,suffix='Categorizando Correos')
            categorias = category(c,emails_total)
            if expectedCategory[contador]==categorias:
                aciertos+=1
            print('-----------------------')
            print(c)
            print('categoría: ', categorias)
            print('categoría esperada: ', expectedCategory[contador])
            print('-----------------------')
            categorias_finales.append(categorias)
            ids_finales.append(i)
            cat_esperada.append(expectedCategory[contador])
            contador+=1
        else:
            break


          
    print('******** Numero de correos categorizados como Otros: ', contador_cat[6])
    print('******** Numero de aciertos: ', aciertos)
    print(len(categorias_finales))
    print(len(cat_esperada))
    WriteExcel(ids_finales,categorias_finales,motivo, cat_esperada)

ReadExcel('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/Input/Correos y formularios.xlsx')




