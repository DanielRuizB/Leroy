import random
import re
from turtle import pos
from xml import sax
import pandas as pd
from pandas import ExcelWriter
import spacy
import string
import math
import datetime
from spacy import tokens
from spacy.matcher import Matcher
from spacy.lang.es import Spanish

def on_match(matches):
    print('Matched!', matches)

nlp = spacy.load('es_core_news_sm') #ss

patterns = [
  [{"LEMMA": {"IN": ["tener", "correcto", "compatible"]}},
  {"POS":"ADP", "OP":"?"},
  {"POS":"DET", "OP":"?"},
  {"LEMMA": {"IN": ["instalación", "instalado", "instalacion"]}}],
]

matcher = Matcher(nlp.vocab)
matcher.add("TEST_PATTERNS", patterns)
doc = nlp(" pongáis en contacto")
past = 0
for token in doc: 
    print(f"{token.text:<10} {token.lemma_:<10} {token.pos_:<10} {token.tag_:<10} {token.dep_:<10} ")
    for morph in token.morph:
        print(token.morph.get("Tense")=="Past")
matches = matcher(doc)
print(past)
on_match(matches)
