from bs4 import BeautifulSoup as bs
import requests

def fetch_data(url, date):
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()  # Tự động raise lỗi nếu status != 200
        return process_data(response.text, date)
    except requests.RequestException as e:
        print(f"Error fetching data from {url}: {e}")
        return None

def process_data(html, date):
    soup = bs(html, "html.parser")
    date = date.strftime("%d/%m/%Y")
    
    title_div = soup.find("div", class_="title")
    if not title_div or date not in title_div.text.strip():
        print("Không tìm thấy ngày trong tiêu đề.")
        return [
            {"giai8": [], "giai7": [], "giai6": [], "giai5": [], 
             "giai4": [], "giai3": [], "giai2": [], "giai1": [], "giaidb": []}
            for _ in range(3)
        ]

    table = soup.find("table", class_="bkqmiennam")
    if not table:
        print("Không tìm thấy bảng kết quả.")
        return None
    
    tinh = [
        {"giai8": [], "giai7": [], "giai6": [], "giai5": [], 
         "giai4": [], "giai3": [], "giai2": [], "giai1": [], "giaidb": []}
        for _ in range(3)
    ]

    province_tables = table.find_all("table", class_="rightcl")
    if not province_tables:
        print("Không tìm thấy bảng kết quả của các tỉnh.")
        return tinh

    for idx, province_table in enumerate(province_tables[:3]):
        province = province_table.find("th")
        if province:
            print(f"Đang xử lý tỉnh: {province.get_text(strip=True)}")
        
        for td in province_table.find_all("td", class_=lambda x: x and x.startswith("giai")):
            prize_category = td.get("class")[0]
            # Lấy số từ <div>, <span>, hoặc văn bản trực tiếp
            numbers = [
                elem.get_text(strip=True)
                for elem in td.find_all(["div"])
                if elem.get_text(strip=True) and not elem.find("span", class_="loading")
            ]
            # Nếu không có <div> hoặc <span>, thử lấy văn bản trực tiếp từ <td>
            if len(numbers) != 0:
                tinh[idx][prize_category].extend(numbers)

    return tinh

# tinh = fetch_data("https://www.minhngoc.net/ket-qua-xo-so/mien-nam/25-04-2025.html")

# print(tinh)