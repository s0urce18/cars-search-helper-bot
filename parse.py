from bs4 import BeautifulSoup
import requests
import os
from currency_converter import CurrencyConverter
import win32com.client

conv = CurrencyConverter()

url = "https://auto.drom.ru/audi/a1/"
page = requests.get(url)

soup = BeautifulSoup(page.text, "lxml")

excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks.Open(u'C:\\Works\\pyBotCarsParse\\cars.xlsx')
ws = wb.ActiveSheet

n = 0

def priceParse(i):
	global n
	try:
		prc = ''.join(str(soup.find_all(class_="css-bhd4b0 e162wx9x0")[i].text)[:str(soup.find_all(class_="css-bhd4b0 e162wx9x0")[i].text).find(" ")].split())
		n+=1
	except:
		prc = 0
	return prc

ruaprc = 0
usaprc = 0

def finishRes():
	global n
	global ruaprc
	global usaprc
	try:
		ruaprc = round(int(int(int(priceParse(0))+int(priceParse(1))+int(priceParse(2))+int(priceParse(3))+int(priceParse(4))+int(priceParse(5))+int(priceParse(6))+int(priceParse(7)))/n))
		usaprc = round(conv.convert(ruaprc, 'RUB', 'USD'))
	except:
		ruaprc = 0
		usaprc = 0
	n = 0

x = None
z = 1
while x == None:
	if ws.Range('E'+str(z)).value == None:
		x = z
		print(x)
	else:
		z += 1

while x < ws.UsedRange.Rows.Count+1:
	url = "https://auto.drom.ru/volvo/"+str(ws.Range('C'+str(x)).value)+"/"
	page = requests.get(url)
	soup = BeautifulSoup(page.text, "lxml")
	finishRes()
	ws.Range('D'+str(x)).value = ruaprc
	ws.Range('E'+str(x)).value = usaprc
	x += 1

ws.Range('A1:E'+str(ws.UsedRange.Rows.Count)).Sort(Key1=ws.Range('E1'), Order1=1, Orientation=1)
wb.Save()
wb.Close()