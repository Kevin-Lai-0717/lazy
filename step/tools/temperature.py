from selenium import webdriver
from selenium.webdriver.common.by import By
import datetime
import random
import time


class Temperature:

    def __init__(self, account, password, mask_off):
        self.account = account
        self.password = password
        self.mask_off = mask_off

    def open_chrome(self):
        self.driver = webdriver.Chrome("chromedriver")


    def connect(self):
        # -----前往TOCC網頁-----
        self.driver.get("https://webapp.cgmh.org.tw/temperature/Home/Login")
        time.sleep(2)

        # -----輸入帳號密碼與點選登入-----
        self.driver.find_element("id", "Account").send_keys(self.account)
        self.driver.find_element("id", "Password").send_keys(self.password)
        time.sleep(1)
        self.driver.find_element(By.CLASS_NAME, "login100-form-btn").click()
        time.sleep(5)

    def my_temperature(self):
        top = 36.7
        bottom = 36.3
        into = random.uniform(top, bottom)
        point = 1

        # -----填入體溫-----
        self.driver.find_element(By.ID, "aTemp").click()
        self.driver.find_element("id", "Temp").send_keys(round(into, point))

        # -----是否與確診者足跡重疊-----
        self.driver.find_element("id", "IsBigAcitvNoN").click()
        time.sleep(5)

        # -----送出-----
        self.driver.find_element(By.XPATH, "//*[@id='Questionn']/div/div/div/div/div/div/a").click()
        time.sleep(5)
        self.driver.close()

        # -----登出-----
        # driver.find_element(By.XPATH, "//*[@id='navbarResponsive']/ul/li[3]/a").click()

    def my_tocc(self):
        # -----點開TOCC並滑到最底-----
        self.driver.find_element(By.ID, "aTocc").click()
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)

        # -----是否有在外脫口罩-----
        if self.mask_off == "n":
            self.driver.find_element(By.XPATH, "// *[ @ id = 'IsBigAcitvNoN']").click()

        elif self.mask_off == "y":
            today = datetime.date.today()
            yesterday = today - datetime.timedelta(days=random.randrange(1, 2, 1))
            yes_m = yesterday.month
            yes_d = yesterday.day
            self.driver.find_element(By.XPATH, "// *[ @ id = 'IsBigAcitvNoNY']").click()
            self.driver.find_element(By.ID, "BIGM").send_keys(yes_m)
            self.driver.find_element(By.ID, "BIGD").send_keys(yes_d)
            time.sleep(1)
            # -----TOCC裡脫口罩的場所如果為餐飲場所 視同沒脫口罩 但我還是用完整XPATH抓到了-----
            self.driver.find_element(By.XPATH,
                                "/html/body/div/section[1]/div/div/div[2]/div/table[1]/tbody/tr[24]/td[2]/div/label[1]/input").click()

        # -----作業地點-----
        self.driver.find_element(By.ID, "OtherOpt").send_keys("研究大樓 3F")

        # -----送出-----
        time.sleep(5)
        self.driver.find_element(By.ID, "btnSave").click()
        time.sleep(5)