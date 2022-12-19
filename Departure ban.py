from bs4 import BeautifulSoup as bs
import pandas as pd
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time

def connectin_init ():
    option = webdriver.ChromeOptions()  # создаем пременую с опциями для вебдрайвера хрома
    option.add_argument("--disable-blink-features=AutomationControlled")
    EXE_PATH = Service('C:\Windows\chromedriver.exe')
    driver = webdriver.Chrome(service=EXE_PATH, options=option)
    driver.get('https://gosuslugi.ru/')
    time.sleep(6)
    try: driver.find_element(By.XPATH, '/html/body/portal-root/header/lib-header/div/div/div[2]/div[2]/div/div/lib-login[1]/lib-button[1]/div/button/span').click()
    except: driver.find_element(By.XPATH, '/html/body/portal-root/header/lib-header/div/div/div[2]/div[2]/div/div').click()
    finally:
        time.sleep(7)

    # Вводим логин и пароль
        driver.find_element(By.ID, "login").click()
        print ("Укажите имя пользователя")
  #  login = input()
        driver.find_element(By.ID, "login").send_keys("9241980666")
        driver.find_element(By.ID, "password").click()
        print("Укажите пароль")
   # password = input()
        driver.find_element(By.ID, "password").send_keys("Meganeedle_7890")
        time.sleep(1)

    # Входим в систему госуслуги
        driver.find_element(By.XPATH, "html/body/esia-root/div/esia-idp/div/div/form/div[4]/button").click()
        time.sleep(3)
        driver.find_element(By.XPATH, "html/body/app-root/main/div/div/div/div/button[2]").click()
        time.sleep(15)
        driver.get("https://www.gosuslugi.ru/404402/1/form")
        time.sleep(15)
        try:
            driver.find_element(By.XPATH,"/html/body/div[1]/div[3]/div[1]/div[2]/div/epgu-form-player/div/div[2]/div/div/div[2]/div/div/div/div/div/div[3]/span").click()
        finally:
            time.sleep(7)
            driver.find_element(By.LINK_TEXT, 'Уточнить адрес').click()
            driver.implicitly_wait(20)


            driver.find_element(By.LINK_TEXT, 'Взыскатель').click()
            time.sleep(10)


           # Не удалять ----------------->
            rez = driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[1]/div[2]/div/epgu-form-player/div/div[2]/div[1]/div/div[1]/form-tree' +
                                          '/epgu-form/div/div/div/div[1]/div[2]/div/form-tree[5]/epgu-form-step/div/div/div/div[4]/form-tree' +
                                          '/epgu-panel/div/div/div/div/div/form-tree/epgu-panel/div/div/div/div/div/form-tree[1]/div/div[1]/div[1]' +
                                          '/div/div[1]')
            rez.click()
            print('Щелкнули по списку, ждем 10 секунд')
            driver.implicitly_wait(30)

            print ('Нашли Input поле')
            rez = driver.find_elements(By.XPATH, "/html/body/div[1]/div[3]/div[1]/div[2]/div/epgu-form-player/div/div[2]/div[1]/div/div[1]/form-tree/epgu-form/" +
                                                 "div/div/div/div[1]/div[2]/div[2]/form-tree[5]/epgu-form-step/div/div/div/div[4]/form-tree/epgu-panel/div/div/" +
                                                 "div/div/div/form-tree/epgu-panel/div/div/div/div/div/form-tree[1]/div/div[1]/div[1]/div/div[1]/div/input")
            print ('Ждем 20 секунд')
            driver.implicitly_wait(30)

            print ('Ищем тэги "li" и накладываем ограничения')

            rez = driver.find_elements(By.TAG_NAME, 'li')
            for k in rez:
                if k.text == 'Хочу наложить ограничения на должника':
                    print(k.text)
                    k.click()
                    driver.implicitly_wait(20)

            rez = driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div[1]/div[2]/div/epgu-form-player/div/div[2]/" +
                                                "div[1]/div/div[1]/form-tree/epgu-form/div/div/div/div[1]/div[2]/div[2]/" +
                                                "form-tree[5]/epgu-form-step/div/div/div/div[4]/form-tree/epgu-panel/div/" +
                                                "div/div/div/div/form-tree/epgu-panel/div/div/div/div/div/form-tree[2]/" +
                                                "div/div[1]/div[1]/div/div[1]/div")
            rez.click()

            print ('Кликнули по уточнению жизненной ситуации')
            driver.implicitly_wait(20)

            rez = driver.find_element(By.XPATH,
                                       "/html/body/div[1]/div[3]/div[1]/div[2]/div/epgu-form-player/div/div[2]/div[1]/div/div[1]/form-tree/epgu-form/div/div/div/div[1]/div[2]/div[2]/form-tree[5]/epgu-form-step/div/div/div/div[4]/form-tree/epgu-panel/div/div/div/div/div/form-tree/epgu-panel/div/div/div/div/div/form-tree[2]/div/div[1]/div[1]/div/div[3]/div/ul/li[2]")
            rez.click()
            rez = driver.find_elements(By.TAG_NAME, 'li')
            for k in rez:
                print(k.text)
                if k.text == 'Хочу запретить должнику выезд из России':
                    k.click()
                    driver.implicitly_wait(20)

            # Блок обработки Исполнительного производства
            #
            # Прочитать список, подставить в поле - "Номер исполнительного производства"
            #

            inpt = driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[1]/div[2]/div/epgu-form-player/div/div[2]/'
                                                 'div[1]/div/div[1]/form-tree/epgu-form/div/div/div/div[1]/div[2]/div[2]/'
                                                 'form-tree[5]/epgu-form-step/div/div/div/div[4]/form-tree/epgu-panel/div/'
                                                 'div/div/div/div/form-tree/epgu-panel/div/div/div/div/div/form-tree[9]/'
                                                 'epgu-field-text/div/div/div/div/div[1]/input')
            inpt.click()
            inpt.send_keys('132456')
            input()

    #
    # driver.find_element(By.CLASS_NAME, "label-model").click()
    # time.sleep(20)
def main():
    connectin_init()

if __name__ == "__main__":
    main()