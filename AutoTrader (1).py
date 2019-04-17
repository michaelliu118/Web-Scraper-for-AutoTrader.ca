import requests
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
import time
import random
import xlwt
from multiprocessing import Pool

class overall():
    def __init__(self, page):
        ua= UserAgent()
        header = {
            'User-Agent': ua.random,
            'Referer': 'https://www.autotrader.ca/bundles/global.7426ac8a.css',
            'cookie': '792d1a4b79feeb28df3a2ee5fd37fe86',
            'Connection': 'keep-alive'}
        proxy = ['69.10.33.81:8080',
                 '134.209.222.103:3128',
                 '198.211.98.159:3128'
                 '147.135.121.131:8080',
                 '115.127.7.162:41500',
                 '147.91.111.133:55508',
                 '12.218.209.130:53281',
                 '109.87.33.2:53381',
                 '38.21.38.6:35458']
        x = random.choice(proxy)
        proxies = {'http': x, 'https': x}
        usepage = (page-1)*15
        url = 'https://wwwa.autotrader.ca/cars/honda/accord/on/?rcp=15&rcs={}&srt=3&prx=-2&prv=Ontario&loc=L4S%201A3&hprc=True&wcp=True&sts=New-Used&inMarket=advancedSearch'.format(usepage)
        try:
            response = requests.get(url, headers=header, proxies=proxies)
            print(proxies)
        except:
            response = requests.get(url, headers=header)
        self.soup = BeautifulSoup(response.text, 'html.parser')
        self.sub = self.soup.find_all('div', class_='col-xs-12 fixed-detail-column')
        self.price_preliminary = self.soup.find_all('div', class_='price-delta')

    def paerser(self):
        models = []
        years = []
        prices =[]
        for i in self.price_preliminary:
            self.price = i.span.text.strip()
            prices.append(self.price)
        for i in self.sub:
            self.content = i.div.h2.a.span.text.strip()
            models.append(str(self.content)[0:4])
            years.append(str(self.content)[5:-1])
        self.info = dict(zip(models, zip(prices, years)))
        print(self.info)
        return self.info

def excel_transfer(x):
    for k, v in x.items():
        sheet.write(row[-1], 0, k)
        sheet.write(row[-1], 1, v[0])
        sheet.write(row[-1], 2, v[1])
        row.append(row[-1]+1)



row = [0]
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('test', cell_overwrite_ok=False)
def main(page):
    time.sleep(random.randint(0,1))
    try:
        a = overall(page)
        return a.paerser()
    except:
        pass
    page += 1

if __name__==  '__main__':
    p = Pool()
    for i in p.map(main, (range(1, 5))):
        excel_transfer(i)
    book.save(r'C:\Users\Xiaokeai\Desktop\AutoTrader Accord1.xls')
