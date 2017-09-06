from selenium import webdriver
import win32gui
import win32con
import win32api
import time
from selenium.webdriver.common.keys import Keys
import selenium.webdriver.support.ui as ui
from bs4 import BeautifulSoup
from PIL import Image
import datetime
    
def get_picture(queueName,P1,P2):
    dt = datetime.datetime.now()
    Ydt = dt + datetime.timedelta(days=-1)
    currentYear = '20'+Ydt.strftime('%y')
    currentMon = Ydt.strftime('%m')
    currentDay = Ydt.strftime('%d')
    DailyFolderName = currentYear+currentMon+currentDay
    #---------BC BL CT-----------
    wait.until(lambda driver:driver.find_element_by_id("from-time"))
    fromtime = driver.find_element_by_id("from-time")
    fromtime.clear()
    current_date = time.time() - (86400*7)
    timeStr = time.strftime("%d %b %Y 00:00",time.localtime(current_date))
    fromtime.send_keys(timeStr)
    time.sleep(1)
    wait.until(lambda driver:driver.find_element_by_id("to-time"))
    totime = driver.find_element_by_id("to-time")
    totime.clear()
    current_date = time.time() - (86400*0)
    timeStr = time.strftime("%d %b %Y 00:00",time.localtime(current_date))
    totime.send_keys(timeStr)
    time.sleep(1)
    #driver.implicitly_wait(5)
    wait.until(lambda driver:driver.find_element_by_name("queueName"))
    queuename = driver.find_element_by_name("queueName")
    queuename.clear()
    queuename.send_keys(queueName)
    time.sleep(1)
    wait.until(lambda driver:driver.find_element_by_name("emsInstance"))
    emsInstance = driver.find_element_by_name("emsInstance")
    emsInstance.clear()
    emsInstance.send_keys('tcp://cosems01v:8001')
    driver.find_element_by_xpath("//*[@id='history-search-criteria']/form/table/tbody/tr[4]/td/input[2]").click()
    time.sleep(15)
    wait.until(lambda driver:driver.find_element_by_xpath("//*[@id='emshis-showchart-1']"))
    driver.find_element_by_xpath("//*[@id='emshis-showchart-1']").click()
    driver.save_screenshot('D:\\DailyReportResouceFiles\\'+DailyFolderName+'\\'+P1)
    time.sleep(10)
    img = Image.open('D:\\DailyReportResouceFiles\\'+DailyFolderName+'\\'+P1)
    region = (10,330,1900,670)
    #裁切图片
    cropImg = img.crop(region)
    #保存裁切后的图片
    cropImg.save('D:\\DailyReportResouceFiles\\'+DailyFolderName+'\\'+P2)
    time.sleep(3)


    
#登录地址
login_url = 'https://www.cargosmart.com/operation/j_spring_security_check'
get_url = 'https://www.cargosmart.com/operation/ems/ems_history.jsf'
driver = webdriver.Firefox()
driver.maximize_window()
wait = ui.WebDriverWait(driver,10)
#driver.set_window_size(1055, 800)  
driver.get(login_url)
username = driver.find_element_by_name("j_username")
username.clear()
username.send_keys('ocuser')
password = driver.find_element_by_name("j_password")
password.clear()
password.send_keys('ocuser')
time.sleep(1)
driver.find_element_by_xpath("//*[@id='submit']").click() 
time.sleep(5)
html=driver.get(get_url)
html=driver.page_source#获取网页的html数据
soup=BeautifulSoup(html,"html.parser")#对html进行解析，如果提示lxml未安装，直接pip install lxml即可    
get_picture("Cargosmart.COSCON.CD.BC.INPUT.PARALLEL","BC1.png","BC2.png")
get_picture("Cargosmart.COSCON.CD.BL.INPUT.PARALLEL","BL1.png","BL2.png")
get_picture("Cargosmart.COSCON.CD.CT.INPUT.PARALLEL","CT1.png","CT2.png")
time.sleep(2)
driver.quit()


