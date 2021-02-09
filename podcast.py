import xml.etree.ElementTree as ET
from openpyxl import Workbook
from urllib.request import urlopen
from matplotlib import pyplot as plt
import numpy as np
import sys

# placeholder variables
export = True
plot = True
bounds = False

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

dates = []
times = []

# move data from XML to workbook
for idx, item in enumerate(items):
	row = []
	for tag in tags:
		if item.find(tag) != None:
			value = item.find(tag).text
			if tag.endswith('duration'):
				value = float(value) / 60
				times.append(value)
			elif tag.endswith('pubDate'):
				dates.append(value)
			row.append(value)
	sheet.append(row)

# save workbook
if export:
	title = root.find('channel').find('title').text
	name = title + '.xlsx'
	wb.save(filename=name)
	print('Wrote to', name)

if plot:
	# plot by episode number
	x = range(1, len(times)+1)
	plt.plot(x, times)

	# add trendlines
	z = np.polyfit(x, times, 1)
	p = np.poly1d(z)
	plt.plot(x, p(x))

	# add labels
	plt.title('"{}" Episode Duration vs Time'.format(title))
	plt.xlabel('Episode Number')
	plt.ylabel('Length (mins)')

	# draw some unnecessary lines for min and max over time
	if bounds:
		maxes = []
		mins = []
		maxXs = []
		minXs = []
		pos = p(1) > p(0)
		for i in range(len(times)):
			val = max(times[0:i+1]) if pos else max(times[i:])
			if times[i] == val:
				maxes.append(val)
				maxXs.append(i+1)
		for i in reversed(range(len(times))):
			val = min(times[i:]) if pos else min(times[0:i+1])
			if times[i] == val:
				mins.append(val)
				minXs.append(i+1)

		# plot min/max lines
		plt.plot(maxXs, maxes)
		plt.plot(minXs, mins)

	# show plot
	plt.show()