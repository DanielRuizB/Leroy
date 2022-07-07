import random
import re
from tkinter import Y
from xml import sax
from numpy import true_divide
import pandas as pd
from pandas import ExcelWriter
import spacy
import string
import math
import datetime
from spacy import tokens
from spacy.matcher import Matcher
from spacy.lang.es import Spanish
from langdetect import detect
#import copy
#from os import remove 
#import requests
#from spacy.lang.en import English
#from spacy.lang.it import Italian
#from spacy.lang.fr import French
#from langdetect import detect
#import pathlib
#import openpyxl 
#from xlrd import open_workbook      # Librería para leer archivos excel .xls como si fuese un libro en bloque

##### CLASES PERSONALIZADAS #####

class Categoria:
    def __init__(self):
        self.palabras = 0 
        self.tfidfs = 0
        self.patrones = 0

class CategoríaEval:
    def __init__(self):
        self.total = 0 
        self.aciertos = 0 
        self.fallos = 0
        self.perdidas = 0

class listaPalabras:
    def __init__(self):
        self.nombre = "" 
        self.palabras = []

#################################################################################

############################ VARIABLES GLOBALES #################################
        
categorias = {}     # Variable global para almacenar los datos necesarios del excel
categorias_finales = []
ids_finales = []
cat_esperada=[]
cat_finales={}
palabrasClave=[]
listaRefs=[]
ltfidfs=[]
#################################################################################

#################### FUNCIONES PARA IMPRIMIR POR TERMINAL #######################
patterns_Asesoramiento_light = [
    [{"LEMMA": {"IN": ["querer", "ir", "desear", "preguntar",  "gustar", "necesitar", "gustariar", "deseariar", "poder", "podriar", "queria", "tener", "podrio"]}},
  {"POS": "ADP", "OP":"?"},
  {"LEMMA": {"IN": ["comprar", "adquirir", "obtener", "enlazar", "ayudar", "poner", "ayudar él", "informar", "confirmar", "solicitar", "consultar", "consultarl", "comunicar yo", "usar", "indicar yo", "decirmir"]}}],
  
]
patterns_Asesoramiento = [
     [{"LEMMA": {"IN": ["querer", "ir", "desear", "buscar", "quiero"]}},
  {"POS": "ADP", "OP":"?"},
  {"LEMMA": {"IN": ["comprar", "adquirir", "obtener", "pintar", "recambio", "pinturas"]}}],

  [{"LOWER": {"IN": ["vendeis", "vendéis", "teneis", "tenéis", "busco", "necesito", "necesitaria"]}},
  {"OP": "*"},
    {"LEMMA": {"IN": ["pintura","césped","cesped","baldas", "soporte"]}}],

  [{"LEMMA": {"IN": ["qué", "que"]}},
  {"LEMMA": {"IN": ["tipo"]}}],

   [{"LEMMA": {"IN": ["me"]}},
  {"LEMMA": {"IN": ["asesoren"]}}],

  [{"LEMMA": {"IN": ["tener", "correcto", "compatible"]}},
  {"POS":"ADP", "OP":"?"},
  {"POS":"DET", "OP":"?"},
  {"LEMMA": {"IN": ["instalación", "instalado", "instalacion"]}}],
    
  [{"LEMMA": {"IN": ["querer", "desear", "gustar", "poder"]}},
  {"POS": "SCONJ"},
  {"POS": "PRON"},
  {"LEMMA": {"IN": ["explicar", "informar", "aconsejar", "consultar","explicarar", "informarar", "aconsejarar", "consultarar", "asesorar"]}}],

  [{"LEMMA": {"IN": ["estar"]}},
  {"LOWER": {"IN": ["interesado", "interesada"]}},
  {"LOWER": {"IN": ["en"]}},
  {"LEMMA": {"IN": ["comprar", "adquirir", "obtener", "saber"]}}], 

   [{"LEMMA": {"IN": ["hacer", "cortar", "fabricarmelar", "fabricármelar"]}},
  {"POS":"ADP"},
    {"POS":"DET", "OP":"?"},
  {"LEMMA": {"IN": ["medida", "cm", "m"]}}],

   [{"LEMMA": {"IN": ["cortar", "fabricarmelar", "fabricármelar"]}},
  {"POS":"ADP"}],

    [{"POS":"ADP"},
    {"POS":"DET", "OP":"?"},
  {"LEMMA": {"IN": ["medida", "cm", "m"]}}],

    [{"LOWER":"no"},
  {"LEMMA": {"IN": ["saber", "aparecer", "salir", "encontrar"]}}],

    [{"LOWER":{"IN": ["que", "si"]}},
    {"LEMMA": {"IN": ["aguantar", "soportar"]}}],

  [{"LEMMA": {"IN": ["necesitar","querer"]}},
      {"POS":"SCONJ"},
  {"LEMMA": {"IN": ["ref", "referencia"]}}],

    [{"LEMMA": {"IN": ["vender"]}},
      {"OP":"*"},
  {"LEMMA": {"IN": ["tipo"]}}],
]

patterns_pasado = [  
    [{"LEMMA": {"IN": ["acabar", "acar"]}}, 
  {"POS": "ADP", "OP":"?"},
  {"LEMMA": {"IN": ["comprar", "adquirir", "obtener", "hacer", "realizar", "llegar", "recibir"]}}],

  [{"LEMMA": {"IN": ["haber", "tras"]}}, 
  {"POS": "ADP", "OP":"?"},
  {"LEMMA": {"IN": ["comprar", "adquirir", "obtener", "hacer", "recibir", "estar", "decir"]}}],

    [{"LOWER": {"IN": ["realicé", "hice", "he"]}}, 
  {"POS": "DET", "OP":"?"},
  {"LEMMA": {"IN": ["pedido", "compra"]}}],

  [{"LOWER": {"IN": ["número", "numero"]}}, 
 {"LOWER": {"IN": ["de"]}},
  {"LEMMA": {"IN": ["pedido"]}}],

     [{"LOWER": {"IN": ["viene", "compré", "compró","compre", "compramos", "adquiri","adquirí", "pagado", "realicé", "comprado", "dispongo", "recibí", "recibido", "recibidas","recibida","llegó", "realizado", "hicimos", "hice", "realice"]}}]
  ]



patterns_Transporte = [
  [{"LEMMA": {"IN": ["entregar", "enviar"]}},
  {"POS": "DET", "OP":"?"},
  {"POS": "ADP", "OP":"?"},
  {"LEMMA": {"IN": ["pedido"]}}],

  [{"LEMMA": {"IN": ["hacer"]}},
  {"LEMMA": {"IN": ["envío"]}}],

[{"LEMMA": {"IN": ["poder"]}},
  {"LEMMA": {"IN": ["enviar", "entregar"]}}],


      [{"LEMMA": {"IN": ["enviar", "entregar", "subir"]}},
  {"POS": "ADP"},
   {"LEMMA": {"IN": ["casa", "domicilio"]}}]
]

patterns_Transporte_prior = [
      [{"LEMMA": {"IN": ["enviar", "entregar", "subir"]}},
  {"POS": "ADP"},
   {"LEMMA": {"IN": ["casa", "domicilio"]}}]

]
patterns_Caracteristicas = [
   [{"LEMMA": {"IN": ["haber","querer"]}},
  {"LEMMA": {"IN": ["comprar", "adquirir"]}}],

   [{"LOWER": {"IN": ["compré","compre"]}},
  ],



     [{"LEMMA": {"IN": ["estar","querer"]}},
  {"LEMMA": {"IN": ["buscar", "confirmar"]}}],
]

patterns_Instalaciones = [
   [{"LEMMA": {"IN": ["querer", "necesitar", "poder"]}},
      {"POS":"DET", "OP":"?"},
  {"LEMMA": {"IN": ["instalar", "montaje"]}}],

     [{"LEMMA": {"IN": ["venir"]}},
      {"POS":"ADP", "OP":"?"},
  {"LEMMA": {"IN": ["instalar", "montar", "montar él"]}}], 

     [{"LEMMA": {"IN": ["para", "incluir", "necesitar", "contratar", "contrato"]}},
      {"POS":"DET", "OP":"?"},
  {"LEMMA": {"IN": ["instalación", "montaje", "instalacion"]}}],


  [{"LEMMA": {"IN": ["contratar", "servicio", "realizar", "precio", "necesitar", "informarmar", "informar"]}},
   {"OP":"*"},
  {"LEMMA": {"IN": ["instalación", "montaje", "colocación", "colocacion", "instalacion"]}}],

    [{"LEMMA": {"IN": ["preguntar"]}},
    {"LOWER": {"IN": ["por"]}},
    {"LOWER": {"IN": ["la", "una", "un", "el"]}},
    {"LEMMA": {"IN": ["instalación", "montaje", "instalacion"]}}],

  [{"LOWER": {"IN": ["opción"]}},
    {"LOWER": {"IN": ["de"]}},
    {"LOWER": {"IN": ["instalación", "montaje"]}}],

]

patterns_Stock_General = [  

      [{"LEMMA": {"IN": ["si"]}},
      {"LEMMA": {"IN": ["estar"]}},
  {"LEMMA": {"IN": ["disponible"]}}],

  [{"LEMMA": {"IN": ["querer", "desear", "gustar", "necesitar"]}}, 
  {"LEMMA": {"IN": ["saber"]}},
  {"POS":"SCONJ"},
  {"LEMMA": {"IN": ["disponer"]}}],

   [{"LEMMA": {"IN": ["tener", "haber"]}},
  {"LOWER": {"IN": ["en"]}},
  {"POS":"?", "LOWER": {"IN": ["alguna", "algún"]}},
  {"LEMMA": {"IN": ["tienda", "almacén"]}}],

  [{"LEMMA": {"IN": ["quedar"]}},
  {"OP":"*"},
  {"LEMMA": {"IN": ["existencia"]}}],

    [{"LEMMA": {"IN": ["estar", "tener"]}},
  {"LEMMA": {"IN": ["en"]}},
  {"LEMMA": {"IN": ["catálogo"]}}],
  
  [{"LEMMA": {"IN": ["volver", ]}},
   {"POS":"ADP", "OP":"?"},
  {"LEMMA": {"IN": ["reponer"]}}],

  [{"LEMMA": {"IN": ["cuando", "cuándo"]}},
   {"OP":"*"},
  {"LEMMA": {"IN": ["reponer"]}}],

    [{"LEMMA": {"IN": ["quedar", "haber", "tener"]}},
  {"LEMMA": {"IN": ["existencia", "disponibilidad"]}}],  #ad

    
]

patterns_Revisar = [
          [{"LEMMA": {"IN": ["servicio"]}},
  {"LEMMA": {"IN": ["técnico", "tecnico"]}}],

      [{"LOWER": {"IN": ["poner", "pongáis", "poneros"]}},
    {"LOWER": {"IN": ["en"]}},
    {"LOWER": {"IN": ["contacto"]}}],

  [{"LEMMA": {"IN": ["no"]}}, 
  {"LOWER": {"IN": ["se", "sé", "puedo", "tengo", "encaja", "cabe", "vale", "valen", "coincide", "llega", "coinciden"]}}],

    [{"LOWER": {"IN": ["solicito", "pido", "quiero"]}},
    {"LOWER": {"IN": ["cancelar", "cancelación"]}}],

    [{"LEMMA": {"IN": ["faltar"]}},
    {"OP":"*"},
    {"LEMMA": {"IN": ["pieza"]}}],

    [{"LOWER": {"IN": ["cuando"]}},
    {"OP":"*"},
    {"LOWER": {"IN": ["llega", "llegar"]}}],

    [{"LOWER": {"IN": ["otro"]}},
    {"LOWER": {"IN": ["color"]}}],

    
    [{"LOWER": {"IN": ["cl", "cliente"]}},
    {"LOWER": {"IN": ["llama"]}}],

     [{"LOWER": {"IN": ["estoy"]}},
    {"LOWER": {"IN": ["esperando"]}}],

        [{"LOWER": {"IN": ["llegado", "recibido"]}}, #sssssssssaa
    {"OP":"*"},
    {"LOWER": {"IN": ["diferente", "distinto"]}}],

        [{"LOWER": {"IN": ["debería", "debe"]}},
        {"LOWER": {"IN": ["ser"]}}],

        
          [{"LOWER": {"IN": ["no"]}},
    {"LOWER": {"IN": ["hay"]}},
    {"LOWER": {"IN": ["transporte"]}}], #transporte

    [{"LOWER": {"IN": ["número", "numero"]}},
    {"LOWER": {"IN": ["de"]}},
    {"LOWER": {"IN": ["pedido"]}}], #transport

     [{"LOWER": {"IN": ["atención"]}},
    {"LOWER": {"IN": ["al"]}},
    {"LOWER": {"IN": ["cliente"]}}],

         [{"LOWER": {"IN": ["en"]}},
    {"LOWER": {"IN": ["vez"]}},
    {"LOWER": {"IN": ["de"]}}],

         [{"LOWER": {"IN": ["que", "qué"]}},
    {"LOWER": {"IN": ["ocurre", "pasa"]}}],

      [{"LOWER": {"IN": ["no"]}},
    {"LOWER": {"IN": ["entra", "funciona", ]}}],

]

patterns_Humano = [
    [{"LOWER": {"IN": ["tomar"]}},
    {"OP":"?", "LOWER": {"IN": ["otras"]}},
    {"LOWER": {"IN": ["medidas"]}}], #transport


  [{"LEMMA": {"IN": ["segunda", "tercera", "cuarta", "2", "3", "4", "varios"]}},
  {"LEMMA": {"IN": ["vez", "petición", "solicitud", "veces", "día", "dia"]}}],

    [{"LEMMA": {"IN": ["mal"]}},
  {"LEMMA": {"IN": ["estado"]}}],



    [{"LEMMA": {"IN": ["no"]}},
   {"LEMMA": {"IN": ["me", "te", "querer", "es", "haber"]}},
    {"OP":"?", "LOWER": {"IN": ["ha"]}},
  {"LEMMA": {"IN": ["decir", "ayudar", "dicen", "saber", "recibir", "dejar", "llegar", "correcto", "llegado"]}}],

   [{"LOWER": {"IN": ["no"]}},
   {"LOWER": {"IN": ["me", "te", "quiero", "es", "ha", "he"]}},
       {"OP":"?", "LOWER": {"IN": ["lo"]}},
    {"OP":"?", "LOWER": {"IN": ["ha", "han"]}},
  {"LOWER": {"IN": ["decir", "ayudan", "dicen", "saber", "recibir", "deja", "llega", "correcto", "llegado", "recibido", "entregado"]}}],


    [{"LOWER": {"IN": ["el", "de"]}},
    {"LOWER": {"IN": ["alta"]}}],

    [{"LOWER": {"IN": ["me"]}},
    {"LOWER": {"IN": ["encuentro"]}}],

]


def detectReference(candidate):
    if ((candidate[0]=='1' or candidate[0]=='7' or candidate[0]=='8') and (len(candidate)==8 and (not ("x" in candidate))) or ("esfp" in candidate)):
        return True
    else:
        return False

#################################################################################

#################### FUNCIONES PARA IMPRIMIR POR TERMINAL #######################

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

def printCategoria(correo,categorias,expected):
    print('-----------------------')
    print(correo)
    print('categoría: ', categorias)
    print('categoría esperada: ', expected)
    print('-----------------------')

def printResultados(cat_finales):
    print('<<<<<<<<<<<<<<<RESULTADOS>>>>>>>>>>>>>>>')
    print()

    for val in cat_finales:
        print('******** Numero de correos categorizados como', val, ': ',  cat_finales[val].total)
        print('******** Numero de aciertos: ',  cat_finales[val].aciertos)
        print('******** Numero de fallos (mal clasificado): ', cat_finales[val].fallos)
        print('******** Numero de fallos (mensajes restantes): ', cat_finales[val].perdidas)
        
        print()

#################################################################################

####################### FUNCIONES PARA CALCULAR TFIDF ###########################

def calcular_tf(correo, palabra):
    numero_palabras = len(correo)
    contador_tf = 0
    for otra_palabra in correo: 
        if otra_palabra.lemma_ == palabra:
            contador_tf = contador_tf + 1

    if(numero_palabras>0):
        return contador_tf/(numero_palabras) 
    else:
        return 0

def calcular_idf(correos_total, palabra):
    contador_idf = 0
    numero_correos= len(correos_total)
    for otro_correo in correos_total:
        for otra_palabra in otro_correo:
            if otra_palabra.lemma_== palabra:
                contador_idf = contador_idf + 1
                break

    if(contador_idf==0):
        contador_idf=1

    return numero_correos/ float(contador_idf) 

def calcular_tfidf(palabra, correo, correos_total):
    # Fórmula del cálculo total tf-idf 
    tfidf = calcular_tf(correo, palabra) * math.log(calcular_idf(correos_total, palabra))
    #valor= contador_tf* tfidf
    return tfidf

#################################################################################

####################### FUNCIONES SOBRE LAS CATEGORÍAS ##########################

def añadirCategoria(valores, categoria, tfidf):
    if not (categoria in valores):
        valores[categoria]=Categoria()
    if (categoria=="Instalaciones"):
        valores[categoria].tfidfs+=tfidf*1
    if (categoria=="Stock"):
        if valores[categoria].patrones==0:
             valores[categoria].tfidfs+=tfidf
        else:
            valores[categoria].tfidfs+=tfidf*(valores[categoria].patrones+1)
    if (categoria=="Asesoramiento experto"):
        if valores[categoria].patrones==0:
             valores[categoria].tfidfs+=tfidf*1
        else:
            valores[categoria].tfidfs+=tfidf*(valores[categoria].patrones)
    if (categoria=="Características generales"):
        if valores[categoria].patrones==0:
             valores[categoria].tfidfs+=tfidf*1
        else:
            valores[categoria].tfidfs+=tfidf*(valores[categoria].patrones+1)
    else:
        valores[categoria].tfidfs+=tfidf
    valores[categoria].palabras+=1

def calcuralValorMax(valores, reference, pasado, pasado_matcher, correo):
    pos=0
    index=0
    max_val=valores["Revisar - Humano"].tfidfs
    for val in valores:
        print(valores[val].tfidfs)
        if (valores[val].tfidfs > max_val):
            max_val = valores[val].tfidfs
            pos=index
        index+=1
    key_list = list(valores.keys())
    print(key_list[pos])
    matcher = Matcher(correo.vocab)
    matcher.add("pasado", patterns_pasado)
    if key_list[pos]!="Asesoramiento experto" and key_list[pos]!="Características generales":
         if len(matcher(correo))>0 or valores["Alerta - Humano"].palabras>0:
            matcher.remove("pasado")
            return "Alerta - Humano"    
    if key_list[pos]=="Stock":
        if len(matcher(correo))>0:
            matcher.remove("pasado")
            return "Revisar - Humano"
        elif len(reference)>0:
            matcher.remove("pasado")
            #print("--------Length: ", len(reference), " -------------------")
            return "Stock - Ref"
        else:
            matcher.remove("pasado")
            return "Stock - noRef"

    if key_list[pos]=="Asesoramiento experto" or key_list[pos]=="Características generales":
        if valores["Alerta - Humano"].palabras>0:
            matcher.remove("pasado")
            return "Alerta - Humano"
        elif len(matcher(correo))>0:
            matcher.remove("pasado")
            return "Características generales"
        else: 
            matcher.remove("pasado")  #pasaod
            return "Asesoramiento experto"

    return key_list[pos]
    
def getCategorias(datwords):
    listCat = []
    for i in datwords.index:
        try:
            listCat.append(listaPalabras())
            listCat[i].nombre = datwords.columns.values[i]
            listCat[i].palabras = datwords.iloc[:,i].to_list()
            cat_finales[listCat[i].nombre]=CategoríaEval()
        except:
            break
 
    return listCat

def guardarCategoria(token,correo,total_correos,valores,listCat,palabrasClave,listPrior):
    tfidf=0
    tfidf=calcular_tfidf(token.lemma_,correo,total_correos) 
    for cat in listPrior:
        if (token.text.lower() in cat.palabras) or (token.lemma_.lower() in cat.palabras):
            palabrasClave.append(token.text)
            #if(len(palabrasClave)>1):
            añadirCategoria(valores, cat.nombre, tfidf*30)
    for cat in listCat:
        if (token.text.lower() in cat.palabras) or (token.lemma_.lower() in cat.palabras):
            palabrasClave.append(token.text)
            #if(len(palabrasClave)>1):
            añadirCategoria(valores, cat.nombre, tfidf)
        else:
            valores["Revisar - Humano"].tfidfs+=tfidf*0.0001   
            valores["Revisar - Humano"].palabras+=1


def inicializarCategoría(listCat, valores):
    for cat in listCat:
        if not (cat.nombre in valores):
            valores[cat.nombre]=Categoria()

def addPatterns(valores, correo):
    matcher = Matcher(correo.vocab)
    matcher.add("Asesoramiento experto", patterns_Asesoramiento)
    valores["Asesoramiento experto"].patrones=len(matcher(correo))
    matcher.remove("Asesoramiento experto")
    matcher.add("Asesoramiento experto", patterns_Asesoramiento_light)
    valores["Asesoramiento experto"].patrones=len(matcher(correo))*0.1
    matcher.remove("Asesoramiento experto")
    matcher.add("Características generales", patterns_Caracteristicas)
    matcher.remove("Características generales")
    matcher.add("Instalaciones", patterns_Instalaciones)
    valores["Instalaciones"].patrones=len(matcher(correo))
    matcher.remove("Instalaciones")
    matcher.add("Transporte", patterns_Transporte)
    valores["Transporte"].patrones=len(matcher(correo))
    matcher.remove("Transporte")
    matcher.add("Stock", patterns_Stock_General)
    valores["Stock"].patrones=len(matcher(correo))
    matcher.remove("Stock")
    matcher.add("Humano", patterns_Humano)
    valores["Alerta - Humano"].patrones=len(matcher(correo))
    matcher.remove("Humano")
    matcher.add("Revisar", patterns_Revisar)
    valores["Revisar - Humano"].patrones=len(matcher(correo))
    matcher.remove("Revisar")

def treatPatterns(valores):
    valores["Stock"].tfidfs+=valores["Stock"].tfidfs*(valores["Stock"].patrones)+(valores["Stock"].patrones)*1
    valores["Asesoramiento experto"].tfidfs+=valores["Asesoramiento experto"].tfidfs+(valores["Asesoramiento experto"].patrones)*1
    valores["Características generales"].tfidfs+=valores["Características generales"].tfidfs+(valores["Características generales"].patrones)*1
    valores["Instalaciones"].tfidfs+=valores["Instalaciones"].tfidfs+(valores["Instalaciones"].patrones)
    valores["Transporte"].tfidfs+=valores["Transporte"].tfidfs+(valores["Transporte"].patrones)*1
    valores["Alerta - Humano"].tfidfs+=valores["Alerta - Humano"].tfidfs+(valores["Alerta - Humano"].patrones)*10
    valores["Revisar - Humano"].tfidfs+=valores["Revisar - Humano"].tfidfs+(valores["Revisar - Humano"].patrones)*5

   
def treatTfidfs(valores, listaidfs):
    listaidfs.append("Características generales: " + str(valores["Características generales"].tfidfs))
    listaidfs.append("Instalaciones: " + str(valores["Instalaciones"].tfidfs))
    listaidfs.append("Transporte: "+ str(valores["Transporte"].tfidfs))
    listaidfs.append("Stock: "+ str(valores["Stock"].tfidfs))
    listaidfs.append("Asesoramiento experto:"+ str(valores["Asesoramiento experto"].tfidfs))
    listaidfs.append("Financiación: "+ str(valores["Financiación"].tfidfs))
    listaidfs.append("Alerta - Humano:"+ str(valores["Alerta - Humano"].tfidfs))
    listaidfs.append("Revisar - Humano: "+ str(valores["Revisar - Humano"].tfidfs))


def calcularCategoria(correo, total_correos, listCat, listPrior):
    valores={}
    valores["Revisar - Humano"]=Categoria()
    listaPalabras=[]
    inicializarCategoría(listCat,valores)
    pasado_m=False
    addPatterns(valores, correo)
    correostring= str(correo) #aaa
    matcher = Matcher(correo.vocab)
    matcher.add("Revisar", patterns_Revisar)

    idioma="es"
    try:
        idioma= detect(correostring.lower())
    except:
       valores["Idioma - Humano"].tfidfs+=valores["Idioma - Humano"].tfidfs+10
    if idioma!="es" and len(matcher(correo))==0: 
       valores["Idioma - Humano"].tfidfs+=valores["Idioma - Humano"].tfidfs+10
    matcher.remove("Revisar")
    reference = False
    pasado = False
    reflist=[]
    listatfidfs=[]
    for token in correo:
        if (not token.is_punct 
            and not token.text == ' ' 
            and not token.like_url 
            and token.lemma_.lower()
            and token.lemma_ != " "):
            guardarCategoria(token,correo,total_correos,valores,listCat,listaPalabras,listPrior)

        ref=False 
        if len(token)==8:   
            reference = detectReference(token.text)
            if reference:
                reflist.append(token.text)
                ref=True
        if "esfp" in token.text:
            reference = detectReference(token.text[token.text.index("esfp")+4:token.text.index("esfp")+12])
            if reference:
                reflist.append(token.text[token.text.index("esfp")+4:token.text.index("esfp")+12])
                ref=True
        elif "ref" in token.text.lower() and len(token.text)==11:
            print(token)
            #reference = detectReference(token.text[token.text.lower().index("ref")+3:token.text.lower().index("ref")+11])
            if reference:
                reflist.append(token.text[token.text.index("ref")+3:token.text.index("ref")+10])
                ref=True
        if len(token.morph.get("Tense"))>0 and pasado==False:   
            pasado = token.morph.get("Tense")[0]=="Past"

    listaRefs.append(reflist)        
    matcher = Matcher(correo.vocab)        
    matcher.add("pasado", patterns_pasado)
    if pasado_m==False and len(matcher(correo))>0:
        pasado_m=True
    matcher.remove("pasado")

    print(pasado_m)
    treatPatterns(valores)
    treatTfidfs(valores, listatfidfs) 
    palabrasClave.append(listaPalabras) 
    ltfidfs.append(listatfidfs)
    return calcuralValorMax(valores, reflist, pasado, pasado_m, correo)


def EvaluarCategoria(categorias, esperado):
    if not (categorias in cat_finales):
        cat_finales[categorias]=CategoríaEval()


    cat_finales[categorias].total+=1
    if esperado==categorias:
        cat_finales[categorias].aciertos+=1
        return 1
    else:
        cat_finales[categorias].fallos+=1
        cat_finales[esperado].perdidas+=1
        return 0

#################################################################################xxssss

####################### FUNCIONES PARA LIMPIAR STRINGS ##########################ssxx


def limpiarString(mensaje, nlp):
    #mensaje_lower = "".join(correos.lower()).split("escribió") #CORREOS
    #mensajecss = re.sub('(?s)<style>(.*?)</style>',' ',mensaje_lower[0]) #CORREOS
    mensaje = str(mensaje)
    men = mensaje.replace(r'<[^>]+>', ' ').strip()      #quitamos etiquetas html
    men = re.sub("\.", " ", men)
    men = re.sub(r'(https|http)?:\/\/(\w|\.|\/|\?|\=|\&|\%)*\b', ' ', men, flags=re.MULTILINE)

    #men = re.sub("\d+", ' ', men)   
    men = re.sub(r"(?<=[a-z])\r?\n"," ", men)
    men  = men.translate(str.maketrans(' ',' ',string.punctuation)) 
    men = " ".join(men.split()) #Eliminamos espacios seguidos
    men = men.strip()    
    return nlp(men)

#################################################################################
    
########################### FUNCIONES SOBRE EL EXCEL ############################
def WriteExcel(id,cat,exp, correos):

    print("Escribimos excel")
    x = datetime.datetime.now()
    do = pd.DataFrame({'ID': id, 'Categoria': cat, 'Esperado': exp, 'Correo':correos, 'Palabras':palabrasClave, 'Referencias': listaRefs, 'TFIDFS':ltfidfs}) 
    writer = ExcelWriter(str(x.strftime('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/Correos Procesados/%d-%m-%Y %H-%M'))+' Categorizados.xlsx')
    do.to_excel(writer, 'Hoja de datos', index=False)
    writer.save()      

def getCorreos(ids,expectedCategory, df, emails_total):
    nlp = spacy.load('es_core_news_sm')
    i=600
    while (i<900):
    #for i in range(300):
        #
        ##
        
        id = df.iloc[i,16]               #Columna 1 -> 'ID'
        correos = df.iloc[i,21]          #Columna 2 -> 'Descripción'
        expectedCategory.append(df.iloc[i,1])
        ids.append(id)
        emails_total.append(limpiarString(correos, nlp))
        i=i+1

    return emails_total   #diccionario con ids de los correos y los correosfaabbxx
    #return random.sample(emails_total, maxCorreos)cvvxxvvcc


def ReadExcel(url_excel):       
    emails_total = []
    ids = []
    info = {}
    expectedCategory=[]
    url_keywords='C:/Users/51232209p/Documents/Libro12.xlsx'
    datwords=pd.read_excel(url_keywords,1)
    priorityWords = pd.read_excel(url_keywords,2)
    listCat = getCategorias(datwords)
    listPrior = getCategorias(priorityWords)

    df = pd.read_excel(url_excel, 0)
    #Eliminamos del dataframa la fila dónde el campo descripción está vacíoss
    #df = df.dropna(subset=["Descripción (Objeto) (Correo electrónico)"])


    info = getCorreos(ids,expectedCategory, df, emails_total)
    print("CATEGORIZAMOS")
    contador = 0
    aciertos = 0
    maxCorreos=len(info)
    for correo in info:#
        printProgressBar(contador,maxCorreos,suffix='Categorizando #Correos')
        categorias = calcularCategoria(correo,info,listCat, listPrior)
        categorias_finales.append(categorias)
        #aciertos+=EvaluarCategoria(categorias, expectedCategory[contador])
        printCategoria(correo,categorias,expectedCategory[contador])#llllllll
        ids_finales.append(contador)
        cat_esperada.append(expectedCategory[contador])
        contador+=1


    printResultados(cat_finales)
    print('******** Numero de aciertos total: ', aciertos)
    WriteExcel(ids,categorias_finales, cat_esperada, info)
    Delete(cat_finales)
    Delete(ids_finales)
    Delete(categorias_finales)
    Delete(cat_esperada)
    Delete(emails_total)
    Delete(datwords)
    Delete(priorityWords)
    Delete(info)
    Delete(patterns_Asesoramiento)
    Delete(patterns_Asesoramiento_light)
    Delete(patterns_Caracteristicas)
    Delete(patterns_Instalaciones)
    Delete(patterns_Transporte)
    Delete(patterns_Stock_General)
    Delete(patterns_Humano)

    
def Delete(cat):
    del cat



#################################################################################ddddsasdssdddss

#################################### MAIN #######################################
 
#ReadExcel('C:/Users/51232209p/Documents/Libro5.xlsx')
#ReadExcel('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/Input/Correos y formularios.xlsx')
#ReadExcel('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/CleanOutput/EjemploReducido.xlsx')
#ReadExcel('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/CleanOutput/Formularios Limpios 2.xlsx')
#ReadExcel('C:/Users/51232209p/Desktop/Libro3.xlsx')
ReadExcel('C:/Users/51232209p/Desktop/Formularios 5.xlsx')
#ReadExcel('C:/Users/51232209p/Desktop/Correos 5.xlsx')
#ReadExcel('C:/Users/51232209p/Desktop/Formularios Red.xlsx')
#ReadExcel('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/Input/Correos y formularios - 300.xlsx')
#ReadExcel('C:/Users/51232209p/Desktop/leroyPython/venvLeroyKW/Input/Corredddos y formularios - 500.xlsx')
#################################################################################
