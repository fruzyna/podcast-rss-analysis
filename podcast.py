import xml.etree.ElementTree as ET
from openpyxl import Workbook

root = ET.parse('rss').getroot()
items = root.findall('.//item')
items.reverse()

wb = Workbook()
sheet = wb.active
sheet.append(['Title', 'Date', 'Length (mins)'])

for idx, item in enumerate(items):
	title = item.find('title').text.split(' - ')[0]
	date = item.find('pubDate').text
	length = float(item.find('{http://www.itunes.com/dtds/podcast-1.0.dtd}duration').text)/60
	sheet.append([title, date, length])

wb.save(filename='durations.xlsx')
