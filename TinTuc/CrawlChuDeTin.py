import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
data = {}
worksheet_dict = {}

def Convert(s):
    t = ""
    for i in range(len(s)-1, -1, -1):
        if s[i] != '/':
            t = t + s[i]
        else:
            break
    t = t + '/'
    t = t[::-1]
    return t

def Crawl_ChuDe(url):
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')    
        uls = soup.find('ul', {'class': 'nav clearfix', 'id': 'menu_nav'})
        if uls:
            li80 = uls.find('li', {'rel': '80'})
            if li80:
                topic = li80.find('a').text
                href = li80.find('a').get('href')
                if len(topic) > 0 and len(href) > 0:
                    if href[0] != '/':
                        href = Convert(href)
                    if topic != 'Hỏi - Đáp':
                        data[href] = topic
                            
            li64 = uls.find('li', {'rel': '64'})
            if li64:
                topic = li64.find('a').text
                href = li64.find('a').get('href')
                if len(topic) > 0 and len(href) > 0:
                    if href[0] != '/':
                        href = Convert(href)
                    if topic != 'Hỏi - Đáp':
                        data[href] = topic
                    
            li84 = uls.find('li', {'rel': '84'})
            if li84:
                topic = li84.find('a').text
                href = li84.find('a').get('href')
                if len(topic) > 0 and len(href) > 0:
                    if href[0] != '/':
                        href = Convert(href)
                    if topic != 'Hỏi - Đáp':
                        data[href] = topic
                
            li16 = uls.find('li', {'rel': '16'})
            if li16:
                ul2 = li16.find('ul', {'class': 'sub-nav2'})
                if ul2:
                    aa = ul2.find_all('a')
                    if aa:
                        for a in aa:
                            topic = a.text
                            href = a.get('href')
                            if len(topic) > 0 and len(href) > 0:
                                if href[0] != '/':
                                    href = Convert(href)
                                if topic != 'Hỏi - Đáp':
                                    data[href] = topic
                            
            li17 = uls.find('li', {'rel': '17'})
            if li17:
                topic = li17.find('a').text
                href = li17.find('a').get('href')
                if len(topic) > 0 and len(href) > 0:
                    if href[0] != '/':
                        href = Convert(href)
                    data[href] = topic
                
            li18 = uls.find('li', {'rel': '18'})
            if li18:
                ul2 = li18.find('ul', {'class': 'sub-nav2'})
                if ul2:
                    aa = ul2.find_all('a')
                    if aa:
                        for a in aa:
                            topic = a.text
                            href = a.get('href')
                            if len(topic) > 0 and len(href) > 0:
                                if href[0] != '/':
                                    href = Convert(href)
                                data[href] = topic
                            
            li15 = uls.find('li', {'rel': '15'})
            if li15:
                ul2 = li15.find('ul', {'class': 'sub-nav2'})
                if ul2:
                    aa = ul2.find_all('a')
                    if aa:
                        for a in aa:
                            topic = a.text
                            href = a.get('href')
                            if len(topic) > 0 and len(href) > 0:
                                if href[0] != '/':
                                    href = Convert(href)
                                if topic != 'Hỏi - Đáp':
                                    data[href] = topic
    return data

def Write_to_Excel(data, file_name):
    wb = Workbook()
    sheet_name = "Chủ Đề"
    worksheet = wb.active
    worksheet.title = sheet_name
    worksheet.append(["Topic", "href"])
    for href, topic in data.items():
        worksheet.append([topic, href])
    wb.save(filename=file_name)
    print('Lưu thành công')

if __name__ == '__main__':
    url = "https://thi.tuyensinh247.com/"
    data = Crawl_ChuDe(url)
    if data:
        Write_to_Excel(data, "Ds_Chu_De.xlsx")

