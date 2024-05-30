from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd


driver = webdriver.Chrome()

page = 1

detail_urls = []
for i in range(0, 100):
    driver.get(f'https://zhkh.club/uk/regions/moskva/page/{page}/')
    page += 1

    try:
        table = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div[5]/table/tbody')
        a_tags = table.find_elements(By.TAG_NAME, 'a')

        for a in a_tags:
            url = a.get_attribute('href')
            detail_urls.append(url)
    except:
        continue


name_list = []
adres_list = []
managers_list = []
phones_list = []
inn_list = []
ogrn_list = []
email_list = []

data = {
    'Наименование': name_list,
    'Адрес': adres_list,
    'Руководитель': managers_list,
    'Телефон(ы)': phones_list,
    'ИНН': inn_list,
    'ОГРН': ogrn_list,
    'E-mail': email_list,
}

for url in detail_urls:
    driver.get(url)
    try:
        table = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table')

        try:
            name = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[1]/td[2]').text
            name_list.append(name or '')
        except:
            name_list.append('')

        try:
            adres = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[2]/td[2]').text
            adres_list.append(adres or '')
        except:
            adres_list.append('')

        try:
            manager = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[3]/td[2]').text
            managers_list.append(manager or '')
        except:
            managers_list.append('')

        try:
            phones = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[4]/td[2]').text
            phones_list.append(phones or '')
        except:
            phones_list.append('')

        try:
            inn = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[5]/td[2]').text
            inn_list.append(inn or '')
        except:
            inn_list.append('')

        try:
            ogrn = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[6]/td[2]').text
            ogrn_list.append(ogrn or '')
        except:
            ogrn_list.append('')

        try:
            email = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[7]/td[2]').text
            email_list.append(email or '')
        except:
            email_list.append('')
    except:
        print(f'[ERROR] Не удалось найти таблицу на странице. Ссылка: {url}')

df = pd.DataFrame(data)
with pd.ExcelWriter('moscow.xlsx') as writer:
    df.to_excel(writer)


'''
Московская область
'''
page = 1

detail_urls = []
for i in range(0, 58):
    driver.get(f'https://zhkh.club/uk/regions/moskovskaia-oblast/page/{page}/')
    page += 1
    try:
        table = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div[5]/table/tbody')
        a_tags = table.find_elements(By.TAG_NAME, 'a')

        for a in a_tags:
            url = a.get_attribute('href')
            detail_urls.append(url)
    except:
        continue


name_list = []
adres_list = []
managers_list = []
phones_list = []
inn_list = []
ogrn_list = []
email_list = []

data = {
    'Наименование': name_list,
    'Адрес': adres_list,
    'Руководитель': managers_list,
    'Телефон(ы)': phones_list,
    'ИНН': inn_list,
    'ОГРН': ogrn_list,
    'E-mail': email_list,
}

for url in detail_urls:
    driver.get(url)
    try:
        table = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table')

        try:
            name = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[1]/td[2]').text
            name_list.append(name or '')
        except:
            name_list.append('')

        try:
            adres = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[2]/td[2]').text
            adres_list.append(adres or '')
        except:
            adres_list.append('')

        try:
            manager = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[3]/td[2]').text
            managers_list.append(manager or '')
        except:
            managers_list.append('')

        try:
            phones = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[4]/td[2]').text
            phones_list.append(phones or '')
        except:
            phones_list.append('')

        try:
            inn = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[5]/td[2]').text
            inn_list.append(inn or '')
        except:
            inn_list.append('')

        try:
            ogrn = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[6]/td[2]').text
            ogrn_list.append(ogrn or '')
        except:
            ogrn_list.append('')

        try:
            email = table.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/main/article/div[2]/div/div[2]/div[1]/table/tbody/tr[7]/td[2]').text
            email_list.append(email or '')
        except:
            email_list.append('')
    except:
        print(f'[ERROR] Не удалось найти таблицу на странице. Ссылка: {url}')

df = pd.DataFrame(data)
with pd.ExcelWriter('MO.xlsx') as writer:
    df.to_excel(writer)