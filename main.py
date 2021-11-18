import requests
import bs4
from openpyxl import Workbook

# download the page
url = 'https://rpa.hybrydoweit.pl'
page = requests.get(url)

#create object 
soup = bs4.BeautifulSoup(page.text, 'html.parser')

#we go to the articles section
articles_section = soup.find(id="articles").find_all('article', class_='rpa-article-card')

#we make a list with the data from the article section
data_export = [[(tag.findChild('a')['title']), ' '.join((tag.find('li').string.split()[1:])), (tag.findChild('a')['href'])] for tag in articles_section]
data_rev = data_export[::-1]

#create excel sheet 
wb = Workbook()
sheet = wb.active
sheet1 = wb.create_sheet("Sheet1")

sheet.cell(row=1,column=1).value = "Tytuł"
sheet.cell(row=1,column=2).value = "Branża/dział"
sheet.cell(row=1,column=3).value = "Link"

sheet1.cell(row=1,column=1).value = "Tytuł"
sheet1.cell(row=1,column=2).value = "Branża/dział"
sheet1.cell(row=1,column=3).value = "Link"

#iterate over the elements and add them to the sheet
for i in range(2,7):
    for k in range(0,3):
        sheet.cell(row=i, column=k+1).value = data_export[i-2][k]

for i in range(2,7):
    for k in range(0,3):
        sheet1.cell(row=i, column=k+1).value = data_rev[i-2][k]

#save sheet
wb.save("excel_file.xlsx")