import requests
import beautifulsoup4 as bs4

s = requests.session()

#headers = {'User-Agent':'Mozilla/5.0',
#           'Referer': 'http://141.51.193.167/simbench/gui/'}
#payload  = {'Simbench_Code': '1-EHVHV-mixed-all-1-no_sw&fomat=powerfactory'}
response  = s.get('https://simbench.de/de/datensaetze/')
print(response)