import os
import sys
import urllib
import openpyxl
import urllib.request
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from pprint import pprint
import re

_basePath = os.path.realpath(os.path.dirname(__file__))
sys.path.append(_basePath)

from Bausteine.xlsxReader import XlsxReader
from Bausteine.Esslinger_Zeitung import harvest_url as esz_liste
from Bausteine.Schw_Tagblatt import harvest_url as schw_tag_liste
from Bausteine.WaiblingerWochblatt import harvest_url as waib_liste
from Bausteine.NuertingerZeitung import harvest_url as nuert_liste
from Bausteine.PforzheimerZeitung import harvest_url as pz_liste

alleAnzeigen=[]
#liste1 = esz_liste()
#liste2 = schw_tag_liste()
#liste3 = waib_liste()
#liste4 = nuert_liste()
liste5 = pz_liste()
#alleAnzeigen = liste1 + liste2 + liste3 + liste4 + liste5
#alleAnzeigen = liste1 + liste2 + liste3 + liste4
alleAnzeigen = liste5



i=0
for Anzeige in alleAnzeigen:
    print(alleAnzeigen[i])
    i=i+1

wb = load_workbook(filename = 'test.xlsx')

dest_filename = 'test.xlsx'

ws1 = wb.active
ws1.title = "WebData"

for i, anzeige in enumerate(alleAnzeigen):

    suchtext = re.search('st', anzeige, flags=0)
    if suchtext is not None:
        #print(dir(suchtext))
        #exit()
        ws1.cell(row=i+1, column=1).value = suchtext.group(0)

    ws1.cell(row=i+1, column=4).value = anzeige

    #print(alleAnzeigen[i])

wb.save(filename=dest_filename)
