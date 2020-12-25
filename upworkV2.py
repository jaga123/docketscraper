from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import re
from requests_html import HTMLSession
import requests
from bs4 import BeautifulSoup as soup
import time
import pandas as pd
from openpyxl.workbook import Workbook

path = "C:\Program Files (x86)\chromedriver.exe"
options = Options()
options.add_argument("--headless")
driver = webdriver.Chrome(path,options=options)
city_names = ["Windsor","Enfield","Hartford","Manchester","Suffield","New Britain","Farmington","Bloomfield","Ellington","Vernon","Somers","Cheshire","Rocky hill",
"Granby","Cromwell","Simsbury","Wethersfield","Newington","Avon","Bristol","Middletown","Glastonbury"]
def get_page_soup(city_name):
    global page_soup
    driver.get("http://civilinquiry.jud.ct.gov/PropertyAddressSearch.aspx")
    search = driver.find_element_by_xpath("//*[@id='ctl00_ContentPlaceHolder1_txtCityTown']")
    search.send_keys(city_name)
    search.send_keys(Keys.RETURN)
    time.sleep(5)
    # Get page source
    page_source = driver.page_source
    page_soup = soup(page_source,'html.parser')
    return page_soup

def get_docket_links(page_soup):
    docket_links = []
    docket_table = page_soup.find(id='ctl00_ContentPlaceHolder1_gvPropertyResults')
    table_rows = docket_table.findAll('tr')
    for row in table_rows:
        try:
            docket_name = (row.find('a').text)
            docket_name = docket_name[:-10]
            docket_year = re.search("20",docket_name)
            if docket_year != None:
                docket_id = (row.find('a').text)
                docket_links.append(f'http://civilinquiry.jud.ct.gov/CaseDetail/PublicCaseDetail.aspx?DocketNo={docket_id}')
        except:
            pass
    return docket_links

def get_docket_data(docket_links):
    for docket_url in docket_links:
        r = requests.get(docket_url)
        content = soup(r.content,'html.parser')
        file_date = content.find(id='ctl00_ContentPlaceHolder1_CaseDetailHeader1_lblFileDate').text
        property_address = content.find(id='ctl00_ContentPlaceHolder1_CaseDetailBasicInfo1_lblPropertyAddress').text
        file_date_formatted = file_date.replace('File Date:','').strip()
        property_address_formatted = property_address.strip()
        party_table = content.find(id='ctl00_ContentPlaceHolder1_CaseDetailParties1_gvParties')
        party_table_rows = party_table.findAll('tr')
        try:
            party_td = party_table.findAll('td')
        except:
            pass
        party_two_name = 'N/A'
        party_two_address = 'N/A'
        for row in party_table_rows:
            party_one = row.find(text='D-01')
            if party_one == 'D-01':
                tables = row.findAll(attrs={"id":True})
                for tid in tables:
                    if re.search('PtyPartyName',str(tid.get('id'))) != None:
                        party_one_name = tid.text
                    if re.search('AppearanceInfo1',str(tid.get('id'))) != None:
                        party_one_address = tid.text
                    elif re.search('NonAppearing',str(tid.get('id'))) != None:
                        party_one_address = 'N/A'

            party_two = row.find(text='D-02')
            if party_two == 'D-02':
                tables = row.findAll(attrs={"id":True})
                for tid in tables:
                    if re.search('PtyPartyName',str(tid.get('id'))) != None:
                        party_two_name = tid.text
                    if re.search('AppearanceInfo1',str(tid.get('id'))) != None:
                        party_two_address = tid.text
                    elif re.search('NonAppearing',str(tid.get('id'))) != None:
                            party_two_address = 'N/A'

        docket_info = {
            'File Date':file_date_formatted,
            'Property Address':property_address_formatted,
            'Defendent 1 Name':party_one_name.strip(),
            'Defendent 1 Address':party_one_address.strip(),
            'Defendent 2 Name':party_two_name.strip(),
            'Defendent 2 Address':party_two_address.strip(),
            'Docket Link':docket_url
        }
        Docket_Data_City.append(docket_info)
        time.sleep(1)

writer = pd.ExcelWriter('output.xlsx')
for city in city_names:
    Docket_Data_City = []
    city_soup = get_page_soup(city)
    city_links = get_docket_links(city_soup)
    get_docket_data(city_links)
    df = pd.DataFrame(Docket_Data_City)
    df.to_excel(writer, index=False, encoding='utf-8-sig', sheet_name=f'{city}')
    time.sleep(5)
writer.save()
driver.close()
driver.quit()
