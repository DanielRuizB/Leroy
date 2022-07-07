import random
import re
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

nlp = spacy.load('es_core_news_sm')
matcher = Matcher(nlp.vocab)
patterns = [
  [{"POS": "VBD"}],
]
matcher.add("TEST_PATTERNS", patterns)
doc = nlp("Quiero comprar una casa. Voy a comprar un regalo.")
matches = matcher(doc)
on_match(matches)