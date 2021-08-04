import sys, pyperclip, requests, smtplib
import bs4, openpyxl, os, re, string
from datetime import date, datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

"""
Scraping Functions
"""
def watchPropertyListingsOnRightmove(productUrl):
    res = requests.get(productUrl)

    print("Request Error: " + str(res.raise_for_status()))

    soup = bs4.BeautifulSoup(res.text, "html.parser")

    try:
        numberOfProperties = len(soup.find_all("div", {"class": "propertyCard"}))
        allPropertyPricesDivs = soup.find_all("span", {"class": "propertyCard-priceValue"})
        allBedroomInfoDivs = soup.find_all("h2", {"class" : "propertyCard-title"})
        allPropertyAddressDivs = soup.find_all("address", {"class": "propertyCard-address"})
        allDateAddedOnDivs = soup.find_all("span", {"class" : "propertyCard-branchSummary-addedOrReduced"})
        allAgentsContactDivs = soup.find_all("a", {"class" : "propertyCard-contactsPhoneNumber"})
        return numberOfProperties, allPropertyPricesDivs, allBedroomInfoDivs, allPropertyAddressDivs, allDateAddedOnDivs, allAgentsContactDivs
    except IndexError:
        print("Selector position may be different on this web-page")

def watchPropertyListingsOnZoopla(productUrl):
    res = requests.get(productUrl)

    print("Request Error: " + str(res.raise_for_status()))

    soup = bs4.BeautifulSoup(res.text, "html.parser")

    try:
        numberOfProperties = len(soup.find_all("a", {"class": "e2uk8e17 css-1g4acnf-StyledLink-Link-StyledLink e33dvwd0"}))
        allPropertyPricesDivs = soup.find_all("p", {"class": "css-1o565rw-Text eczcs4p0"})
        allBedroomInfoDivs = soup.find_all("h2", {"class" : "css-vthwmi-Heading2-StyledAddress e2uk8e13"})
        allPropertyAddressDivs = soup.find_all("p", {"class": "css-nwapgq-Text eczcs4p0"})
        allDateAddedOnDivs = soup.find_all("span", {"data-testid" : "date-published"})
        allAgentsContactDivs = soup.find_all("a", {"data-testid" : "agent-phone-number"})
        return numberOfProperties, allPropertyPricesDivs, allBedroomInfoDivs, allPropertyAddressDivs, allDateAddedOnDivs, allAgentsContactDivs
    except IndexError:
        print("Selector position may be different on this web-page")

rmUrl = "https://www.rightmove.co.uk/property-to-rent/find.html?searchType=RENT&locationIdentifier=REGION%5E87494&insId=1&radius=0.5&minPrice=&maxPrice=1750&minBedrooms=2&maxBedrooms=2&displayPropertyType=&maxDaysSinceAdded=&sortByPriceDescending=&_includeLetAgreed=on&primaryDisplayPropertyType=&secondaryDisplayPropertyType=&oldDisplayPropertyType=&oldPrimaryDisplayPropertyType=&letType=&letFurnishType=&houseFlatShare="
numberOfPropertiesR, allPropertyPricesR, allBedroomInfoDivsR, allPropertyAddressDivsR, allDateAddedOnDivsR, allAgentsContactDivsR = watchPropertyListingsOnRightmove(rmUrl)
zpUrl = "https://www.zoopla.co.uk/to-rent/property/london/bloomsbury/?beds_max=2&beds_min=2&price_frequency=per_month&price_max=1750&q=Bloomsbury%2C%20London&results_sort=newest_listings&search_source=home"
numberOfPropertiesZ, allPropertyPricesZ, allBedroomInfoDivsZ, allPropertyAddressDivsZ, allDateAddedOnDivsZ, allAgentsContactDivsZ = watchPropertyListingsOnZoopla(zpUrl)

"""
Writing to excel files
"""
workbook = openpyxl.Workbook()
rightMoveSheet = workbook['Sheet']
rightMoveSheet.title = "RightMove"
zooplaSheet = workbook.create_sheet("Zoopla")

"""
Writing data from Rightmove
"""
columnCellNames = list(string.ascii_uppercase)
columnsToUse = columnCellNames[:5]

currentCell = ""
currentColumnCount = 0
maxCol = len(columnsToUse)
fieldToEnterR = [allPropertyPricesR, allBedroomInfoDivsR, allPropertyAddressDivsR, allDateAddedOnDivsR, allAgentsContactDivsR]
for col in columnsToUse:
    currentField = fieldToEnterR[currentColumnCount]
    for row in range(0, numberOfPropertiesR):
        cell = col + str(row + 1)
        rightMoveSheet[cell] = currentField[row].text
    currentColumnCount += 1


"""
Writing data from Zoopla
"""
currentCell = ""
currentColumnCount = 0
maxCol = len(columnsToUse)
fieldToEnterZ = [allPropertyPricesZ, allBedroomInfoDivsZ, allPropertyAddressDivsZ, allDateAddedOnDivsZ, allAgentsContactDivsZ]
for col in columnsToUse:
    currentField = fieldToEnterZ[currentColumnCount]
    for row in range(0, numberOfPropertiesZ):
        cell = col + str(row + 1)
        zooplaSheet[cell] = currentField[row].text
    currentColumnCount += 1

"""
saving the file with the date and time 
"""
today = date.today()
now = datetime.now()
current_time = now.strftime("%H-%M-%S")
xlfileName = (str(today) + "_" + str(current_time) + '.xlsx')
workbook.save(xlfileName)

"""
SENDING THE EMAIL
"""
file = xlfileName
nameFile = str(today) + "_" + str(current_time)
username='jwww@gmail.com'
password='htydjyjyjtukjyukyukykketykjey'
send_from = 'jwww@gmail.com'
send_to = 'jwww@gmail.com'
Cc = 'recipient'
msg = MIMEMultipart()
msg['From'] = send_from
msg['To'] = send_to
msg['Cc'] = Cc
msg['Subject'] = 'Daily Property Scrape Results'
server = smtplib.SMTP("smtp.gmail.com", 587)
fp = open(file, 'rb')
part = MIMEBase('application','vnd.ms-excel')
part.set_payload(fp.read())
fp.close()
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment', filename=nameFile)
msg.attach(part)
smtp = smtplib.SMTP("smtp.gmail.com", 587)
smtp.ehlo()
smtp.starttls()
smtp.login(username,password)
smtp.sendmail(send_from, send_to.split(',') + msg['Cc'].split(','), msg.as_string())
smtp.quit()
print("Done")
