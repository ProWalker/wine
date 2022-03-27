import collections
import datetime
import os

import pandas

from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape
from dotenv import load_dotenv

load_dotenv()

data_file_path = os.getenv('DATA_FILE_PATH')
worksheet_name = os.getenv('WORKSHEET_NAME')

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')
excel_data_df = pandas.read_excel(data_file_path, worksheet_name, na_filter=False)
wines = collections.defaultdict(list)

for dictionary in excel_data_df.to_dict(orient='records'):
    key = dictionary.pop('Категория')
    wines[key].append(dictionary)

now = datetime.datetime.now()
years_delta = now.year - 1920

if years_delta % 100 in range(11, 20):
    company_age = str(years_delta) + ' лет'
elif years_delta % 10 in [0, 5, 6, 7, 8, 9]:
    company_age = str(years_delta) + ' лет'
elif years_delta % 10 == 1:
    company_age = str(years_delta) + ' год'
elif years_delta % 10 in [2, 3, 4]:
    company_age = str(years_delta) + ' года'

rendered_page = template.render(company_age=company_age, wines=wines)

with open('index.html', 'w', encoding="utf-8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
