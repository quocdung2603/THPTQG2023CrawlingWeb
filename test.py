import requests
from bs4 import BeautifulSoup
ListPost = []

def GetPost(urll, getLink):

    url1 = urll+getLink
    # url1 = 'https://dantri.com.vn/lao-dong-viec-lam/doanh-nghiep-dong-loat-doi-ngay-lam-viec-de-dip-le-304-nghi-keo-dai-20240406100931932.htm'
    res = requests.get(url1)
    if res.status_code == 200:
        soupContent = BeautifulSoup(res.text, 'html.parser')
        CheckPost = soupContent.find('article', class_='e-magazine')
        if (CheckPost is not None):
            return
        ContentPost = soupContent.find('div', class_='singular-content')
        DictContent = {}
        DictContent["Author"] = GetAuthorPost(soupContent)
        DictContent["Title"] = GetTitle(soupContent)
        DictContent["Category"] = getCategory(soupContent)
        DictContent["Content"] = []
        if (ContentPost is not None):
            for child in ContentPost.children:
                if (child.name == 'p'):
                    dic = {}
                    dic["text"] = child.text
                    dic["image"] = ""
                    DictContent["Content"].append(dic)
                elif (child.name == 'figure'):
                    img_tag = child.find('img')
                    if (img_tag is None or img_tag['data-src'] is None):
                        continue
                    dic = {}
                    dic["text"] = ""
                    dic["image"] = img_tag['data-src']
                    DictContent["Content"].append(dic)
        ListPost.append(DictContent)

def GetAuthorPost(data):
    # tên tác giả và thời gian đăng bài
    Author = {}
    Author['Name'] = ''
    Author['Time'] = ''
    check = data.find('div', class_='author-name')
    if (check is not None):
        Author['Name'] = check.text
    else:
        Author['Name'] = "Not found"
    check = data.find('time', class_='author-time')
    if (check is not None):
        Author['Time'] = check.text
    else:
        Author['Time'] = "Not found"
    return Author

def GetTitle(data):
    # Tiêu đề bài viết
    Title = ''
    Title = data.find('h1', class_='title-page').text
    return Title

def getCategory(data):
    # Lấy danh mục bài viết
    Category = ''
    check = data.find('ul')
    if (check is not None):
        Category = check.find('li').text
    return Category

# Gửi yêu cầu GET đến trang web
url = "https://dantri.com.vn"
response = requests.get(url)

if response.status_code == 200:
    # Sử dụng BeautifulSoup để phân tích HTML
    soup = BeautifulSoup(response.text, 'html.parser')
    step = 0
    data = soup.find_all('article', class_="article-item")
    cnt = 0
    for item in data:
        cnt += 1
        GetPost(url, item.a.get('href'))  # hàm lấy nội dung
        if (cnt == 100):
            break
    print(ListPost)
else:
    print("Failed to retrieve the webpage")
