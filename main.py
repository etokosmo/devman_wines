import collections
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape
from pytils import numeral

EXCEL_CATALOG = 'wine3.xlsx'
EXCEL_SHEET = 'Лист1'
CREATION_YEAR = 1920


def get_winery_age() -> str:
    """Get year since creation of the company"""
    today = datetime.date.today()
    winery_age = today.year - CREATION_YEAR
    winery_age_with_caption = numeral.get_plural(winery_age, "год, года, лет")
    return winery_age_with_caption


def get_wines_catalog() -> collections.defaultdict:
    """Get products catalog from Excel"""
    products = pandas.read_excel(EXCEL_CATALOG, sheet_name=EXCEL_SHEET, na_values=['N/A', 'NA'],
                                 keep_default_na=False).to_dict(orient='records')
    products_by_categories = collections.defaultdict(list)
    for product in products:
        products_by_categories[product['Категория']].append(product)
    return products_by_categories


def main():
    """Start server"""
    products_catalog = get_wines_catalog()

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        winery_age_with_caption=get_winery_age(),
        products=products_catalog,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
