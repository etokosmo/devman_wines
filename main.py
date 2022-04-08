import collections
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape
import pandas
from pytils import numeral

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')
catalog = 'wine3.xlsx'
sheet = 'Лист1'
creation_year = 1920


def get_since_creation_year() -> str:
    """Get year since creation of the company"""
    today = datetime.date.today()
    years_since_creation_number = today.year - creation_year
    years_since_creation_value = numeral.get_plural(years_since_creation_number, "год, года, лет")
    return years_since_creation_value


def get_wines_catalog() -> collections.defaultdict:
    """Get wines catalog from Excel"""
    excel_data_df = pandas.read_excel(catalog, sheet_name=sheet, na_values=['N/A', 'NA'], keep_default_na=False)
    headers = excel_data_df.columns.tolist()
    count_wines = len(excel_data_df[headers[1]].tolist())
    wines = collections.defaultdict(list)

    for i in range(count_wines):
        category = excel_data_df[headers[0]][i]
        wines[category].append(
            {
                'title': excel_data_df[headers[1]][i],
                'sort': excel_data_df[headers[2]][i],
                'price': f'{excel_data_df[headers[3]][i]} р.',
                'image': f'images/{excel_data_df[headers[4]][i]}',
                'sale': excel_data_df[headers[5]][i]
            }
        )
    return wines


wines_catalog = get_wines_catalog()

rendered_page = template.render(
    year_since_creation=get_since_creation_year(),
    wines=wines_catalog,
    categories=wines_catalog.keys(),
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
