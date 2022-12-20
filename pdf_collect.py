import time

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from tqdm import tqdm


def wait():
    input("Devam etmek için 'Enter' tuşuna basın.")


def scroll(driver):
    pause_time = 3
    last_height = driver.execute_script("return document.body.scrollHeight")

    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight - 1500);")

        time.sleep(pause_time)

        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height


browser = webdriver.Chrome('./binaries/drivers/win/chromedriver.exe')

browser.get('https://investmentpolicy.unctad.org/international-investment-agreements')

wait()
scroll(browser)

print('Veriler çıkarılıyor.')
data_row = []
t_rows = browser.find_elements(By.CSS_SELECTOR, 'table.table.ajax > tbody > tr')
for t_row in tqdm(t_rows):
    r_id = t_row.find_element(By.TAG_NAME, 'th').text
    r_tds = t_row.find_elements(By.TAG_NAME, 'td')

    r_title = r_tds[1].text
    r_type = r_tds[2].text
    r_status = r_tds[3].text
    r_parties = r_tds[4].text
    r_date = r_tds[5].text
    r_date2 = r_tds[6].text
    r_texts = r_tds[8].find_elements(By.TAG_NAME, 'a')

    if not r_type == 'BITs':
        continue

    pdf_langs = {}
    for r_text in r_texts:
        pdf_langs[r_text.text] = r_text.get_attribute('href')

    if len(pdf_langs):
        if pdf_langs.get('en'):
            r_lang = 'en'
            r_link = pdf_langs.get('en')
        elif pdf_langs.get('tr'):
            r_lang = 'tr'
            r_link = pdf_langs.get('tr')
        else:
            first_key = list(pdf_langs.keys())[0]
            r_lang = first_key
            r_link = pdf_langs.get(first_key)
    else:
        r_lang = ''
        r_link = ''

    data_row.append([
        r_title,
        r_type,
        r_status,
        r_parties,
        r_date,
        r_date2,
        r_lang,
        r_link
    ])

df = pd.DataFrame(data_row, columns=[
    'Short Title',
    'Type',
    'Status',
    'Parties',
    'Date of Signature',
    'Date of Entry Into Force',
    'PDF Lang',
    'PDF Url'
])
df.to_excel('data.xlsx')
