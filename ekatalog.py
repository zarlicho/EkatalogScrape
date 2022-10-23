from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import requests # request img from web
import shutil # save img locally
import urllib.request
import xlsxwriter
import pandas as pd
df = pd.read_excel('katalogData.xlsx')
outWorkbook = xlsxwriter.Workbook("katalogData.xlsx")
outSheet = outWorkbook.add_worksheet()
outSheet.write("A1","Title")
outSheet.write("B1","Etalase Produk")
outSheet.write("C1","Tanggal mulai")
outSheet.write("D1","Tanggal akhir")
outSheet.write("E1","Detail")

opt = Options()
opt.add_argument('--headless')
opt.add_argument('--disable-gpu')
driver =webdriver.Chrome(ChromeDriverManager().install(), chrome_options=opt)
# driver =webdriver.Chrome(ChromeDriverManager().install())
actions = ActionChains(driver)

row = len(df.index)
nilai = 0
driver.get("https://e-katalog.lkpp.go.id/pengumuman")
time.sleep(5)
for j in range(10):
    for fr in range(1,12):
        nilai+=1
        katalog = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[3]/div[1]/div[1]/div[2]/div/div[1]/div/div[{rawr}]".format(rawr=fr))))
        # print(katalog.text)
        spliter = katalog.text
        Produk = spliter.split("Etalase Produk :")
        tglM = spliter.split("Tanggal Mulai:")
        tglA = spliter.split("Tanggal Akhir:")
        a = tglM[1]
        b = tglA[1]
        c = Produk[1]
        x = b.split("Tanggal Akhir:")[0]
        y = x.split("Lokal")[0]
        z = y.split("Detail")[0]
        print(Produk[0])
        outSheet.write(nilai+1,0,Produk[0])
        outSheet.write(nilai+1,1,c.split("Tanggal Mulai:")[0])
        outSheet.write(nilai+1,2,a.split("Tanggal Akhir:")[0])
        outSheet.write(nilai+1,3,z)
        outSheet.write(nilai+1,4,c.split("Detail")[1])
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[3]/div[1]/div[1]/div[2]/div/div[2]/div/div[1]/ul/li[12]/a"))).click()
    time.sleep(3)
print(outWorkbook)
outWorkbook.close()
