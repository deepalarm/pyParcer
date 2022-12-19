from bs4 import BeautifulSoup as bs
import pandas as pd
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time
import pdb
from datetime import datetime, date, timedelta
from bs4 import BeautifulSoup as bs
import random



def get_web():
    pdb.set_trace()
    option = webdriver.FirefoxOptions()
    option.set_preference("dom.webdriver.enabled", False)
    # option.set_preference("general.useragent.override","Mozilla/5.0 (")
    ran = random.randint(6, 15)  # рэндом что бы не спалиться
    today_day = datetime.today().weekday()  # сегодняшний день
    print (today_day)
    start = 1  # старт счетчика если понедельник
    check = True  # выход из цикла

    file = open("index.html", "a")
    # Service('C:/path to file/geckodriver-v0.30.0-win64/geckodriver.exe')
    driver = webdriver.Firefox(
        executable_path='C:/path to file/geckodriver-v0.30.0-win64/geckodriver.exe',
        options=option)

    # # ____________обманывем скрипты по обнарудению Selenium___________#
    # driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    #     "source": """
    #           const newProto = navigator.__proto__
    #           delete newProto.webdriver
    #           navigator.__proto__ = newProto
    #           """
    # })
    # # ____________________________________________________________________#

    driver.get('https://kad.arbitr.ru/')
    time.sleep(7)

    if today_day == 0:
        while check:
            try:
                for i in range(start, 4):
                    new_date = (datetime.today().date() - timedelta(i))

                    date = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div[1]/dl/dd/div[5]/label[1]/input")

                    for y in range(0, 8):
                        date.send_keys(Keys.BACKSPACE)
                    date.send_keys(new_date.strftime("%d%m%Y"))

                    date1 = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div[1]/dl/dd/div[5]/label[2]/input")

                    for z in range(0, 8):
                        date1.send_keys(Keys.BACKSPACE)
                    date1.send_keys(new_date.strftime("%d%m%Y"))

                    el = driver.find_element(By.CLASS_NAME, 'bankruptcy')
                    el.click()
                    time.sleep(6)

                    file.write(driver.page_source)

                    for a in range(3, 13):
                        time.sleep(ran)
                        el = driver.find_element(By.XPATH, f'/html/body/div[2]/div/div[3]/div/ul/li[{a}]/a').click()
                        time.sleep(2)
                        file.write(driver.page_source)

                    for b in range(4, 15):
                        time.sleep(ran)
                        el = driver.find_element(By.XPATH, f'/html/body/div[2]/div/div[3]/div/ul/li[{b}]/a').click()
                        time.sleep(2)
                        file.write(driver.page_source)

                    for c in range(4, 15):
                        time.sleep(ran)
                        el = driver.find_element(By.XPATH, f'/html/body/div[2]/div/div[3]/div/ul/li[{c}]/a').click()
                        time.sleep(2)
                        file.write(driver.page_source)

                    for d in range(4, 14):
                        time.sleep(ran)
                        el = driver.find_element(By.XPATH, f'/html/body/div[2]/div/div[3]/div/ul/li[{d}]/a').click()
                        time.sleep(2)
                        file.write(driver.page_source)
                check = False

            except:
                print(f"На дату {new_date} ни чего нет")
                start = start + 1  # продолжаем после ловли исключения

    else:
        try:

            new_date = (datetime.today().date() - timedelta(1))
            convert_new_date = str(new_date.strftime("%d%m%Y"))
            date = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div[1]/dl/dd/div[5]/label[1]/input")

            webdriver.ActionChains(driver).click_and_hold(date).perform()
            webdriver.ActionChains(driver).release(date).perform()

            for i in convert_new_date:
                webdriver.ActionChains(driver).send_keys(i).perform()

            date1 = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div[1]/dl/dd/div[5]/label[2]/input")

            webdriver.ActionChains(driver).click_and_hold(date1).perform()
            webdriver.ActionChains(driver).release(date1).perform()

            for i in convert_new_date:
                webdriver.ActionChains(driver).send_keys(i).perform()

            el = driver.find_element(By.CLASS_NAME, 'bankruptcy active')
            el.click()
            time.sleep(6)

            file.write(driver.page_source)

            for a in range(3, 13):
                time.sleep(ran)
                el = driver.find_element(By.XPATH, f'/html/body/div[2]/div/div[3]/div/ul/li[{a}]/a').click()
                time.sleep(2)
                file.write(driver.page_source)

            for b in range(4, 15):
                time.sleep(ran)
                el = driver.find_element(By.XPATH, f'/html/body/div[2]/div/div[3]/div/ul/li[{b}]/a').click()
                time.sleep(2)
                file.write(driver.page_source)

            for c in range(4, 15):
                time.sleep(ran)
                el = driver.find_element(By.XPATH, f'/html/body/div[2]/div/div[3]/div/ul/li[{c}]/a').click()
                time.sleep(2)
                file.write(driver.page_source)

            for d in range(4, 14):
                time.sleep(ran)
                el = driver.find_element(By.XPATH, f'/html/body/div[2]/div/div[3]/div/ul/li[{d}]/a').click()
                time.sleep(2)
                file.write(driver.page_source)

        except:
            print(f"На дату {new_date} ни чего нет")

    driver.close()
    driver.quit()
    file.close()


def pasrsing():  # парсинг файла index.html
    with open("index.html") as file:
        src = file.read()

    soup = bs(src, "html.parser")

    td = []
    table2 = soup.find_all("table", id="b-cases")

    for i in range(0, len(table2)):
        for y in table2[i].find_all("td", class_="respondent"):
            fe = y.find("strong")
            td.append(fe)

    for item in td:
        if item is None:
            td.remove(item)

    new_list = []
    [new_list.append(item) for item in td if item not in new_list]

    df = pd.DataFrame(new_list)
    df.to_excel("FIO1.xlsx")
    file.close()


if __name__ == "__main__":
    get_web()  # сохраняем все в файл Index.html. После сохранения в файл закоментить
    # pasrsing()  расскоментрировать когда страницы сохранены.
