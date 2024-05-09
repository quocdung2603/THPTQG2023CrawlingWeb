from openpyxl import Workbook
import requests
from bs4 import BeautifulSoup

def crawl_tinh(url):
    dataTinh = {}
    response = requests.get(url)
    if response.status_code == 200 : 
        soup = BeautifulSoup(response.content, 'html.parser')
        tables = soup.find_all('table', {'class':'table table-bordered'})
        
        for table in tables:
            rows = table.find_all('tr')
            for row in rows[1:]:  # Bỏ qua hàng đầu tiên chứa tiêu đề cột
                cells = row.find_all('td')
                if len(cells) >= 2:  # Đảm bảo có đủ thông tin tên và mã tỉnh/thành phố
                    ten_tinh = cells[0].get_text().strip()
                    ma_tinh = cells[1].get_text().strip()
                    if str(ten_tinh)!="" and str(ma_tinh)!="" and len(str(ma_tinh)) == 2:
                        dataTinh[ma_tinh] = ten_tinh
        
        dataTinh = dict(sorted(dataTinh.items()))
        return dataTinh
    else:
        print('Failed to retrieve data')
        return None

def write_to_excel(data, file_name):
    wb = Workbook()
    ws = wb.active
    ws.append(['Mã tỉnh', 'Tên tỉnh'])  # Tiêu đề cột

    for ma_tinh, ten_tinh in data.items():
        ws.append([ma_tinh, ten_tinh])

    wb.save(file_name)
    print(f'Dữ liệu đã được lưu vào file: {file_name}')


if __name__ == '__main__':
    url = "https://www.vietjack.com/thong-tin-tuyen-sinh/"
    dataTinh = crawl_tinh(url)
    if dataTinh:
        write_to_excel(dataTinh, 'data_tinh.xlsx')
