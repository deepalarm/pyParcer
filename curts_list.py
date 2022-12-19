from bs4 import BeautifulSoup as bs
import pandas as pd
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time
from datetime import datetime, date, timedelta
from tkcalendar import Calendar
import random
import tkinter as tk

option = webdriver.ChromeOptions()  # создаем пременую с опциями для вебдрайвера хрома
option.add_argument("--disable-blink-features=AutomationControlled")  # прячемся от определения вебдрайвера
ran = random.randint(6, 15)  # рэндом что бы не спалиться
check = True  # выход из цикла
file = open("index.html", "w", encoding="utf-8")  # открываем файл на дозапись и сохраняем в него страницы
start = 0  # счетчик для продолжения цикла с судами
EXE_PATH = Service('C:\Windows\chromedriver.exe')
driver = webdriver.Chrome(service=EXE_PATH, options=option)
driver.get('https://kad.arbitr.ru/')
time.sleep(7)
file.write(driver.page_source)

with open("index.html", encoding="utf-8") as file:
    src = file.read()

soup = bs(src, "html.parser")

for line in soup.find_all('select', id='Courts'):
    courts_file = open('courts.txt', 'w', encoding='utf-8')
    print(line.text)
    courts_file.write(line.text)


