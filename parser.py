import csv
import re

import pandas as pd
import requests
from bs4 import BeautifulSoup


phone_number_pattern = r'\+?\d\s?\(\d+\)\s?\d{3}-\d{2}-\d{2}'
email_pattern = r'\b[A-Za-z0-9._%±]+@[A-Za-z0-9.-]+.[A-Z|a-z]{2,}\b'
contacts_pattern = r'контакты'
file = open('./дичь.csv', 'w', encoding='utf-8-sig', newline='')
writer = csv.writer(file, delimiter=';')
writer.writerow(['Отрасль', 'Компания', 'Телефон', 'E-mail', 'URL', 'Код ответа сервера'])
xl = pd.ExcelFile('./Организации сайта ФЦК ВСЕ.xlsx')
sheets = {sheet_name: xl.parse(sheet_name) for sheet_name in xl.sheet_names}

for sheet_name, df in sheets.items():
    print(f'Обрабатываем лист: {sheet_name}')
    for index, row in df.iterrows():
        company_name = str(row.iloc[1]).replace('\n', '')
        company_url = str(row.iloc[12]).replace('"', '').replace('\n', '')
        try:
            response = requests.get(company_url, timeout=15)
            response.encoding = 'utf-8'
            text = response.text
            email = set(re.findall(email_pattern, text))
            phone_number = set(re.findall(phone_number_pattern, text))
            soup = BeautifulSoup(response.text, 'lxml')
            link = soup.find('a', text=re.findall(contacts_pattern, text, re.IGNORECASE))
            if link:
                link = soup.find('a', text=re.findall(contacts_pattern, text, re.IGNORECASE))['href']
                contact_r = requests.get(link)
                contact_r.encoding = 'utf-8'
                phone_number = phone_number.update(set(
                    re.findall(phone_number_pattern, contact_r.text)))
                email = email.update(set(
                    re.findall(email_pattern, contact_r.text)))
                company_url = contact_r
                    
            flatten = sheet_name, company_name, phone_number, email, company_url, response.status_code
            print(flatten)
            writer.writerow(flatten)

        except:
            flatten = sheet_name, company_name, 'Не удалось подключиться', 'Не удалось подключиться', f'Сайт не отвечает ({company_url})', 'Код ответа сервера: {response.status_code}'
            writer.writerow(flatten)
            continue

file.close()