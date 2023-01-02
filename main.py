import requests
import openpyxl
from datetime import date
from html.parser import HTMLParser

# from html.parser import HTMLParser
s = requests.get('http://30rates.com/btc-to-usd-forecast-today-dollar-to-bitcoin').text


class MyHTMLParser(HTMLParser):

    # Initializing lists
    lsStartTags = list()
    lsEndTags = list()
    lsStartEndTags = list()
    lsdata = list()

    # HTML Parser Methods
    def handle_starttag(self, start_tag, attrs):
        self.lsStartTags.append(start_tag)

    def handle_endtag(self, end_tag):
        self.lsEndTags.append(end_tag)

    def handle_startendtag(self, start_end_tag, attrs):
        self.lsStartEndTags.append(start_end_tag)

    def handle_data(self, data):
        self.lsdata.append(data)

    def error(self, message):
        pass


parser = MyHTMLParser()

parser.feed(s)

Price_index = parser.lsdata.index('Price')

# last_index = parser.lsdata.index('13891')
#
# print(last_index) -> 394

parser.lsdata[:] = parser.lsdata[:395]

parser.lsdata[:] = parser.lsdata[Price_index + 1:]

parser.lsdata[:] = [value for value in parser.lsdata if value != '$']

parser.lsdata[:] = [value for value in parser.lsdata if value != ' \n\r\n \n']

wb = openpyxl.load_workbook('/Users/Reza/Desktop/all/UNIVRSITY/python/btc/datacatch.xlsx')
ws = wb.active

counter = 0

start_date = date(2018, 5, 6)
today = date.today()

section = (today - start_date).days

for col in range(section * 27 + 2, section * 27 + int(len(parser.lsdata) / 5) + 2):
    for i in ['A', 'B', 'C', 'D', 'E']:
        string = i + str(col)
        ws[string] = parser.lsdata[counter]
        counter += 1

wb.save("/Users/Reza/Desktop/all/UNIVRSITY/python/btc/datacatch.xlsx")
