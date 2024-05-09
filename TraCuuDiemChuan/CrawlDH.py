from openpyxl import Workbook
import requests
from bs4 import BeautifulSoup

def Crawl_TenDH(url):
    dataDH = []
    response = requests.get(url)
    if response.status_code == 200: 
        soup = BeautifulSoup(response.content, 'html.parser')
        uls = soup.find('ul', {'class':'list_style', 'id':'benchmarking'})
        if uls:
            lis = uls.find_all('li')
            for li in lis:
                a = li.find('a')
                if a:
                    title = a.get('title')
                    href = a.get('href')
                    strong = a.find('strong',{'class':'clblue2'})
                    if strong:
                        tvt = strong.text
                        #print(title,tvt,href)
                        dataDH.append((title,tvt,href))
    return dataDH

def Write_to_excel(data, filename):
    wb = Workbook()
    ws = wb.active
    ws.append(['Tên đại học','Tên viết tắt','Tên tỉnh'])
    for title,tvt,href in data:
        ws.append([title,tvt,href])
        print("Đã lưu trường",title)
    wb.save(filename)

if __name__ == '__main__':
    url = "https://diemthi.tuyensinh247.com/"
    data = Crawl_TenDH(url)
    Write_to_excel(data, 'data_TenDH.xlsx')
