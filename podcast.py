import xml.etree.ElementTree as ET
from openpyxl import Workbook
from urllib.request import urlopen
import sys

# download podcast from RSS URL
if len(sys.argv) < 2:
	exit('No URL provided')
url = sys.argv[1]
rss = urlopen(url).read().decode('utf-8')

# parse and sort feed
root = ET.fromstring(rss)
items = root.findall('.//item')
items.reverse()

# create workbook
wb = Workbook()
sheet = wb.active
sheet.append(['Title', 'Date', 'Length (mins)'])

# move data from XML to workbook
for idx, item in enumerate(items):
	title = item.find('title').text.split(' - ')[0]
	date = item.find('pubDate').text
	length = float(item.find('{http://www.itunes.com/dtds/podcast-1.0.dtd}duration').text)/60
	sheet.append([title, date, length])

# save workbook
title = root.find('channel').find('title').text
name = title + '.xlsx'
wb.save(filename=name)
print('Wrote to', name)
