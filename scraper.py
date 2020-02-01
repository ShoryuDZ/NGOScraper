import requests
import urllib.request
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from googlesearch import search

NGOWorkBook = load_workbook(filename = 'NGOWorkBook.xlsx')
NGOSheet = NGOWorkBook['NGOSheet']

isContent = True
rowNumber = 2

while isContent:
    cellValue = NGOSheet['A' + str(rowNumber)].value
    if cellValue:
        for url in search(cellValue, stop=1):
            print(url)
        
        rowNumber += 1
    else:
        isContent = False
    