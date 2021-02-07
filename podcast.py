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

# get unique tags
tags = [e.tag for e in items[0].iter() if e.tag != 'item']
names = [t if '}' not in t else t[t.rindex('}')+1:] for t in tags]

# create workbook
wb = Workbook()
sheet = wb.active
sheet.append(names)

# move data from XML to workbook
for idx, item in enumerate(items):
	row = []
	for tag in tags:
		if item.find(tag) != None:
			value = item.find(tag).text
			if tag.endswith('duration'):
				value = float(value) / 60
			row.append(value)
	sheet.append(row)

# save workbook
title = root.find('channel').find('title').text
name = title + '.xlsx'
wb.save(filename=name)
print('Wrote to', name)
