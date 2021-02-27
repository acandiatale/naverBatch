import requests
from bs4 import BeautifulSoup
import excelController

r = requests.get('https://finance.naver.com/sise/sise_group.nhn?type=upjong')

soup = BeautifulSoup(r.text, 'html.parser')
items = soup.select('.type_1 > tr > td > a')
values = soup.select('.type_1 > tr > td > span')

itemlist = dict()

for item in zip(items, values):
    itemlist[(item[0].text)] = item[1].text.replace('\n', '').replace('\t', '')
    
excelController.invoke(itemlist)