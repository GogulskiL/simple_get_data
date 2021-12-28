from datetime import date
import requests
import bs4
from openpyxl import Workbook, workbook

# download the page
url = 'https://rpa.hybrydoweit.pl'
title = "Tytuł"
industry = "Branża/dział"
link = "link"


def get_data(url):
    return requests.get(url)


def create_soup(url):
    data = get_data(url)
    return bs4.BeautifulSoup(data.text, 'html.parser')


def get_article_data(url):
    articles_section = create_soup(url).find(
        id="articles").find_all('article', class_='rpa-article-card')
    return [[(tag.findChild('a')['title']), ' '.join((tag.find('li').string.split()[
        1:])), (tag.findChild('a')['href'])] for tag in articles_section]


def get_article_data_reverse(url):
    return get_article_data(url)[::-1]


# create excel sheet
# wb = Workbook()
# sheet = wb.active
# sheet1 = wb.create_sheet("Sheet1")

# sheet.cell(row=1, column=1).value = "Tytuł"
# sheet.cell(row=1, column=2).value = "Branża/dział"
# sheet.cell(row=1, column=3).value = "Link"

# sheet1.cell(row=1, column=1).value = "Tytuł"
# sheet1.cell(row=1, column=2).value = "Branża/dział"
# sheet1.cell(row=1, column=3).value = "Link"

# data_export = get_article_data(url)
# data_rev = get_article_data_reverse(url)
# # iterate over the elements and add them to the sheet
# for i in range(2, 7):
#     for k in range(0, 3):
#         sheet.cell(row=i, column=k+1).value = data_export[i-2][k]

# for i in range(2, 7):
#     for k in range(0, 3):
#         sheet1.cell(row=i, column=k+1).value = data_rev[i-2][k]

# # save sheet
# wb.save("excel_file.xlsx")
def create_excel_file(name_cell_1, name_cell_2, name_cell_3):
    wb = Workbook()
    sheet = wb.active
    sheet1 = wb.create_sheet("Sheet1")
    sheet.cell(row=1, column=1).value = name_cell_1
    sheet.cell(row=1, column=2).value = name_cell_2
    sheet.cell(row=1, column=3).value = name_cell_3

    sheet1.cell(row=1, column=1).value = name_cell_1
    sheet1.cell(row=1, column=2).value = name_cell_2
    sheet1.cell(row=1, column=3).value = name_cell_3
    wb.save("data_file.xlsx")

def fill_excel_file(excel_file, url):
    wb = create_excel_file(title, industry, link)
    data = get_article_data(url)
    data_rev = get_article_data_reverse(url)
    for i in range(2, 7):
        for k in range(0, 3):
            sheet.cell(row=i, column=k+1).value = data[i-2][k]

    for i in range(2, 7):
        for k in range(0, 3):
            sheet1.cell(row=i, column=k+1).value = data_rev[i-2][k]
    