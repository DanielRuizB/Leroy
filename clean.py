import datetime
import re
import pandas as pd
from pandas import ExcelWriter
import string

def WriteExcel(id,cat,exp):

    print("Escribimos excel")
    x = datetime.datetime.now()
    do = pd.DataFrame({'ID': id, 'Categoria': cat, 'Mensaje': exp}) 
    #writer = ExcelWriter(str(x.strftime('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/CleanOutput/%d-%m-%Y %H-%M'))+' Correos Limpios.xlsx')
    #writer = ExcelWriter(str('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/CleanOutput/Correos Limpios.xlsx'))
    writer = ExcelWriter(str('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/CleanOutput/Formularios Limpios 2.xlsx'))
    do.to_excel(writer, 'Hoja de datos', index=False)
    writer.save() 

def ReadExcel(url_excel):       

    emails_total = []
    ids = []
    expectedCategory=[]

    df = pd.read_excel(url_excel, 0)
    #Eliminamos del dataframa la fila dónde el campo descripción está vacío
    #df = df.dropna(subset=["Descripción (Objeto) (Correo electrónico)"])
    id=0
    for i in df.index:
        
        id = df.iloc[i,16]               #Columna 1 -> 'ID'ss
        correos = df.iloc[i,21]          #CORREOS
        expectedCategory.append(df.iloc[i,16])         #CORREOS
        #correos = df.iloc[i,16]          #FORMULARIOS
        #expectedCategory.append(df.iloc[i,14])         #FORMULARIOS
        
        #men_split=correos.lower().split("escribió") #CORREOS
        men=correos.split("Comentarios:") #FORMULARIOS
        print(men[0])
        #if len(men_split_form)>1:
            #men=men_split_form[1]
        #else:
            #men=men_split_form[0]

        #men = re.sub('(?s)<style>(.*?)</style>',' ',men) #CORREOS
        #men = re.sub('<.*?>|&([a-z0-9]+|#[0-9]{1,6}|#x[0-9a-f]{1,6});',' ',men) #FORMULARIOS
        men = men[1].replace(r'<[^>]+>', ' ').strip()      #quitamos etiquetas html
        men = re.sub("\.", " ", men)
        men = re.sub(r'(https|http)?:\/\/(\w|\.|\/|\?|\=|\&|\%)*\b', ' ', men, flags=re.MULTILINE)

        #men = re.sub("\d+", ' ', men)   
        men = re.sub(r"(?<=[a-z])\r?\n"," ", men)
        en = re.sub("\."," ", men)
        men  = men.translate(str.maketrans(' ',' ',string.punctuation)) 
        men = " ".join(men.split()) #Eliminamos espacios seguidos
        men = men.strip()

        ids.append(id)
        emails_total.append(men)
    print("LIMPIAMOS")
    WriteExcel(ids, expectedCategory, emails_total)
#url_excel=('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/CleanInput/Correos y formularios.xlsx')
#url_excel=('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/CleanInput/Correos y formularios.xlsx')
#url_excel=('C:/Users/51232209p/Desktop/Bot de formularios_Tipología Motivo Financiación.xlsx')
#url_excel=('C:/Users/51232209p/Desktop/Correos-22-3.xlsx')
url_excel=('C:/Users/51232209p/Desktop/Formularios 5.xlsx')
#url_excel=('C:/Users/51232209p/Desktop/Formularios Red.xlsx')
#url_excel=('C:/Users/51232209p/Desktop/LibroCorreos.xlsx')
#url_excel=('C:/Users/51232209p/Desktop/Bot de correos_Tipología Motivo Financiación.xlsx') #Formularios pasados por clean.py
#url_excel=('C:/Users/51232209p/Desktop/Bot de formularios_Tipología Motivo Financiación.xlsx') #Formularios pasados por clean.py
ReadExcel(url_excel)#ggg
