import requests
import urllib.request
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from googlesearch import search

NGOWorkBook = load_workbook(filename = 'NGOWorkBook.xlsx')
NGOSheet = NGOWorkBook['NGOSheet']

isContent = True
rowNumber = int(input("Enter first row of data: "))

def linkFinder(url, deliveryList, organisation, mode):
    try:
        response = requests.get(url)

        soup = BeautifulSoup(response.text, "html.parser")
        Links = soup.body.find_all('a')

        for Link in Links:
            LinkString = str(Link).lower()
            if mode == "Email" && LinkString.find('mailto') != -1:
                deliveryList.append(Link['href'])
            elif mode == "ContactURLs" && LinkString.find('contact') != -1:
                if Link['href'].startswith('http'):
                    deliveryList.append(Link['href'])
                elif Link['href'].startswith('/'):
                    deliveryList.append(url[:-1] + Link['href'])
                else:
                    deliveryList.append(url + Link['href'])
    except:
        print("Error on {0} mode of {1}".format(mode, organisation))

def IFLGoogle(query):
    urlGenerateds = search(cellValue, stop=1)
    for urlGenerated in urlGenerateds:
        return urlGenerated
    
def emailFormatter(emailList):
    emailString = ''
    emailList = set(emailList)
    for email in emailList:
        if email == emailList[-1]:
            emailString += email[7:]
            break
        else:
            emailString += (email[7:] + "; ")
    return emailString
    
while isContent:
    cellValue = NGOSheet['A' + str(rowNumber)].value
    if cellValue:
        URL = IFLGoogle(cellvalue)
        contactURLs = []
        
        linkFinder(URL, contactURLs, cellValue, "ContactURLs")
           
        emails = []

        for contactURL in contactURLS:
            linkFinder(contactURL, emails, cellValue, "Email")
        
        emailString = emailFormatter(emails)

        NGOSheet['C' + str(rowNumber)].value = emailString
        NGOWorkBook.save('NGOWorkBook.xlsx')
            
        rowNumber += 1

    else:
        isContent = False
    