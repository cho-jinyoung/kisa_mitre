from bs4 import BeautifulSoup
from selenium import webdriver
import requests
import re
import csv
import openpyxl
import time

wb=openpyxl.Workbook()  # 작업할 workbook생성
sheet=wb.active         # 작업할 workbook내 sheet활성화

sheet.append(["tactic", "technique","ID", "설명", "subtechnique", "ID", "설명", "map attack ID(tactic name)" ])

# 웹페이지 접속
driver=webdriver.Chrome("./chromedriver.exe")
driver2=webdriver.Chrome("./chromedriver.exe")
driver.get("https://d3fend.mitre.org")  
time.sleep(3)           # selenium이 웹페이지를 읽어올 시간을 줌
html=driver.page_source
soup=BeautifulSoup(html)

# tactic 요소 이름 읽음
css_selector="#mwrap > main > section > div > div:nth-child(1) > div.tree-branch-fork.tree-level-0.svelte-ldyryq > a"
tactic=driver.find_element_by_css_selector(css_selector)
#print(tactic.text)

# technique 요소 이름 읽음
css_selector="#mwrap > main > section > div > div:nth-child(1) > div.tree-trunk.svelte-ldyryq > div:nth-child(1) > div.tree-branch-fork.tree-level-1.svelte-ldyryq > a"
technique=driver.find_element_by_css_selector(css_selector)
#print(technique.text)


for ink in soup.find_all('div', {'class':'tree-leaf svelte-arxkma'}):
    subtac_url=ink.a['href']
    subtac_title=ink.string
    #print(subtac_title+" >>sub_tectic_url= https://d3fend.mitre.org/"+subtac_url)
    
    driver2.get("https://d3fend.mitre.org/"+subtac_url)
    time.sleep(1)
    html2=driver2.page_source
    soup2=BeautifulSoup(html, 'html.parser')

    tac_add="#sapper > main > main > section > div.modal.svelte-1xyq1x2 > section:nth-child(2) > div > div:nth-child(1) > div"
    tac_id=driver2.find_element_by_css_selector(tac_add+" > span > span:nth-child(2) > a")
    tac_def=driver2.find_element_by_css_selector(tac_add+" > div > p:nth-child(1)")
    #print(tac_id.text, tac_def.text, "\n")  

css_selector="#mwrap > main > section > div > div:nth-child(1) > div.tree-trunk.svelte-ldyryq > div"
subtec=driver.find_elements_by_css_selector(css_selector)
for i, item in enumerate(subtec):
    #print(item.text)

    item.find_element_by_css_selector('div > a').click()
    item_text=item.get_attribute('herf')
    #print(item_text)

driver.quit()
wb.save("D3fend.xlsx")
