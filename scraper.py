import requests
import urllib.request
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from googlesearch import search

NGOWorkBook = load_workbook(filename = 'NGOWorkBook.xlsx')
NGOSheet = NGOWorkBook['NGOSheet']

isContent = True
rowNumber = 17

while isContent:
    cellValue = NGOSheet['A' + str(rowNumber)].value
    if cellValue:
        try:
            urlGenerateds = search(cellValue, stop=1)
            for urlGenerated in urlGenerateds:
                url = urlGenerated
            response = requests.get(url)

            soup = BeautifulSoup(response.text, "html.parser")
            Links = soup.body.find_all('a')

            contactURLS = []

            for Link in Links:
                LinkString = str(Link).lower()
                if LinkString.find('contact') != -1:
                    if Link['href'].startswith('http'):
                        contactURLS.append(Link['href'])
                    elif Link['href'].startswith('/'):
                        contactURLS.append(url[:-1] + Link['href'])
                    else:
                        contactURLS.append(url + Link['href'])

            emails = []
            emailString = ''

            for contactURL in contactURLS:
                print(contactURL)
                contactResponse = requests.get(contactURL)
                soup = BeautifulSoup(contactResponse.text, "html.parser")
                contactLinks = soup.body.find_all('a')

                for Link in contactLinks:
                    LinkString = str(Link).lower()
                    if LinkString.find('mailto') != -1:
                        emails.append(Link['href'])

                for email in emails:
                    if email == emails[-1]:
                        emailString += email[7:]
                        break
                    else:
                        emailString += (email[7:] + "; ")

            NGOSheet['C' + str(rowNumber)].value = emailString
            NGOWorkBook.save('NGOWorkBook.xlsx')

        except:
            print("Error for: " + cellValue)
            
        rowNumber += 1

    else:
        isContent = False
    