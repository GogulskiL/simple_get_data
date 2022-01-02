import requests
import bs4
from openpyxl import Workbook

# download the page
url = 'https://rpa.hybrydoweit.pl'
title = "Title"
industry = "Industry"
link = "Link"


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


def create_excel_file(excel_file_name, number_sheet, name_cell_1, name_cell_2, name_cell_3):
    wb = Workbook(excel_file_name)
    for sheet in range(number_sheet):
        sheet = wb.create_sheet(f"Sheet {sheet}")
        sheet.append([str(name_cell_1), str(name_cell_2), str(name_cell_3)])
    wb.save(f"{excel_file_name}.xlsx")


create_excel_file("a", 4, title, industry, link)
