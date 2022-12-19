# pyParcer
Python parser for structuring and processing requests in FSSP.
To work, you need to install the COM component for the data base, in this example, the COM component for 1C is added.


You need change #location directory# to parcer data 

Depature ban.py

    # Вводим логин и пароль
        driver.find_element(By.ID, "login").click()
        print ("Укажите имя пользователя")
  #  login = input()
        driver.find_element(By.ID, "login").send_keys("LOGIN")
        driver.find_element(By.ID, "password").click()
        print("Укажите пароль")
   # password = input()
        driver.find_element(By.ID, "password").send_keys("PASSWRD")
        time.sleep(1)

369.py

# Инициализурием COM подключение
conn_str = "Srvr='SERVER EITH DATA BASE';Ref='_______ ';Usr='TYPE USER NAME HERE';Pwd='TYPE PASSWORD HERE';"
pythoncom.CoInitialize()
V83 = win32com.client.Dispatch("V83.COMConnector").Connect(conn_str)
current_dir = str(new_fold_dir)
dir_to_copy = ('C:\Output destination')

_mail.py 

file = open("index.html", "a")
    # Service('C:/Users/baryshnikov_va.MEGASAKH/Downloads/geckodriver-v0.30.0-win64/geckodriver.exe')
    driver = webdriver.Firefox(
        executable_path='C:/Users/baryshnikov_va.MEGASAKH/Downloads/geckodriver-v0.30.0-win64/geckodriver.exe',
        options=option)
        
        
