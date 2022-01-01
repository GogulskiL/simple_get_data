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


def create_excel_file(excel_file_name, number_sheet):
    wb = Workbook(excel_file_name)
    for sheet in range(number_sheet):
        sheet = wb.create_sheet(f"Sheet{sheet}")
    wb.save(f"{excel_file_name}.xlsx")


def create_column( name ,excel_file, number_sheet):
    excel_file = Workbook(create_excel_file(excel_file, number_sheet))
    sheet = excel_file.active()
    sheet.cell(row = 1, column=1).value = name
