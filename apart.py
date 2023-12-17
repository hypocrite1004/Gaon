import requests
import json
import pandas as pd
import win32com.client as win32
import folium
from folium.plugins import MarkerCluster
import json
from collections import defaultdict
import xml.etree.ElementTree as elemTree
from bs4 import BeautifulSoup

def get_data(local_code, date):
    url = 'http://openapi.molit.go.kr/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptTradeDev'
    all_data = []
    for lc in local_code:
        params = {
            'serviceKey': 'APIKEY',   # 공공데이터 포털에서 받은 인증키
            'pageNo': '1',
            'numOfRows': '1000',
            'LAWD_CD': lc,
            'DEAL_YMD': date
        }

        response = requests.get(url, params=params)
        soup = BeautifulSoup(response.text, 'lxml-xml')
        items = soup.find_all('item')

        data = []

        for i in items:
            row = {
                '거래금액': i.find('거래금액').text,
                '건축년도': i.find('건축년도').text,
                '년': i.find('년').text,
                '월': i.find('월').text,
                '일': i.find('일').text,
                '아파트': i.find('아파트').text,
                '전용면적': i.find('전용면적').text,
                '층': i.find('층').text,
                '등기일자': i.find('등기일자').text,
                '법정동': i.find('법정동').text,
                '지번' : i.find('지번').text,
                '도로명': i.find('도로명').text if i.find('도로명') else 'N/A',
                '도로명건물본번호코드': i.find('도로명건물본번호코드').text,
                '도로명건물부번호코드': i.find('도로명건물부번호코드').text,
                '해제사유발생일': i.find('해제사유발생일').text,
                '해제여부': i.find('해제여부').text
            }
            data.append(row)
            
        all_data.extend(data)
        print(lc + ' 완료')
    
    df = pd.DataFrame(all_data)
    df.to_excel('apart.xlsx', index=False)
    print('엑셀 생성 완료')

def addr_to_lat_lon(addr):
    url = 'https://dapi.kakao.com/v2/local/search/address.xml?query={address}'.format(address=addr)
    headers = {"Authorization": "KakaoAK " + 'APIKEY'}    # KaKao Rest API 키
    
    response = requests.get(url, headers=headers)
    # print(response.text)
    tree = elemTree.fromstring(response.text)
    document = tree.find('documents')        
    lat = document[0].find('y').text
    lon = document[0].find('x').text
    return lat, lon

def get_geo(file):
    df = pd.read_excel(file)
    
    latitude = []
    longitude = []
    
    for a, b, c, d, e in zip(df['법정동'], df['지번'], df['도로명'], df['도로명건물본번호코드'], df['도로명건물부번호코드']):
        addr = '부산광역시' + ' ' + a + ' ' + str(b)
        try:
            lat, lon = addr_to_lat_lon(addr)
            latitude.append(lat)
            longitude.append(lon)
        except TypeError:
            addr = c + ' ' + str(d) + '-' + str(e)
            lat, lon = addr_to_lat_lon(addr)
            latitude.append(lat)
            longitude.append(lon)
            
        print(a, b)
    
    df['위도'] = latitude
    df['경도'] = longitude

    with pd.ExcelWriter(file, mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, index=False)
    
    # df.to_excel('apart_test.xlsx', index=False)
    print('위도 경도 추가 완료')

def make_m():
    with open('Kor_Map.json', mode='r', encoding='utf-8') as f:
        geo = json.load(f)

    df = pd.read_excel('apart_test.xlsx')

    apart = df['아파트']
    lat = df['위도']
    lon = df['경도']
    price = df['거래금액']
    footage = df['전용면적']

    m = folium.Map(location=[35.238318470813, 129.081156957989], zoom_start=13)

    info_by_location = defaultdict(list)
    for apart, lat, lon, price, foot in zip(apart, lat, lon, price, footage):
        info_by_location[(lat, lon)].append((apart, price, str(foot)))
    
    marker_cluster = MarkerCluster().add_to(m)

    for location, infos in info_by_location.items():
        lat, lon = location
        info_string = '<br>'.join(f'아파트: {apart}, 전용면적: {str(foot)}, 거래가격: {price}' for apart, price, foot in infos)
        iframe = folium.IFrame('<pre>' + info_string + '</pre>')
        popup = folium.Popup(iframe, min_width=500, max_width=500)
        folium.Marker(
            location=[lat, lon], 
            icon=folium.Icon(color='blue'), 
            popup=popup
        ).add_to(marker_cluster)

    folium.GeoJson(geo, name='geojson').add_to(m)

    m.save('apart.html')
    print('지도 생성 완료')

def excel_form(file):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    load_wb = excel.Workbooks.Open(file)
    
    for sh in load_wb.Sheets:
        ws = load_wb.Worksheets(sh.Name)
        ws.Columns.AutoFit()
        ws.Rows.AutoFit()
        ws.Cells.HorizontalAlignment = win32.constants.xlCenter

    load_wb.Save()
    excel.Application.Quit()
    print('엑셀 서식 맞춤 완료')

local_code = ['26440', '26410', '26710', '26290', '26170', '26260', '26230', '26320', '26530', '26380', '26140', '26500', '26470', '26200', '26110', '26350']
date = '202311'

get_data(local_code, date)
get_geo('apart.xlsx')
make_m()
excel_form('apart.xlsx')

'''
부산광역시	법정동코드
강서구	26440
금정구	26410
기장군	26710
남구	26290
동구	26170
동래구	26260
부산진구	26230
북구	26320
사상구	26530
사하구	26380
서구	26140
수영구	26500
연제구	26470
영도구	26200
중구	26110
해운대구	26350
'''
