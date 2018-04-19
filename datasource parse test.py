from bs4 import BeautifulSoup as bs

file = "C:/Users/rlynch/Downloads/Support Case Queue.twb"

with open(file) as soupstock:
	soup = bs(soupstock, "lxml")

datasources = soup.datasources.find_all('datasource')

for datasource in datasources:
	print(datasource.get('caption'))
	columns = []
	for column in datasource.find_all('column'):
		# print(column.prettify())
		# print(column.get('name'))
		# print(column.get('datatype'))
		columns.append((column.get('name').replace('[','').replace(']',''), column.get('datatype')))
	[print("    " + column[0] + " | " + column[1].upper()) for column in columns]
