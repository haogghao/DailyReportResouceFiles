#encoding:UTF-8
import time
import datetime
from selenium import webdriver
import selenium.webdriver.support.ui as ui
from selenium.webdriver.common.keys import Keys
import datetime

##获取每日的文件夹
dt = datetime.datetime.now()
Ydt = dt + datetime.timedelta(days=-1)
currentYear = '20'+Ydt.strftime('%y')
currentMon = Ydt.strftime('%m')
currentDay = Ydt.strftime('%d')
DailyFolderName = currentYear+currentMon+currentDay
##设置谷歌浏览器下载配置
options = webdriver.ChromeOptions()
prefs = {'profile.default_content_settings.popups':0,'download.default_directory':'D:\\DailyReportResouceFiles\\'+DailyFolderName}
options.add_experimental_option('prefs',prefs)
driver = webdriver.Chrome(chrome_options=options)
wait = ui.WebDriverWait(driver,5)
time.sleep(1)
##打开CS2-ACZ-COSCON-PROD的网址
url = 'http://artdc.cargosmart.com:8000/zh-CN/app/CS4_E2E_MON/CS2_Front_End_Process_Monitoring_ART_Report?form.INPUT_APPLICATION_NAME=CS2-ACZ-COSCON-PROD&form.INPUT_MODULE_NAME=*&form.INPUT_SERVICE_NAME=*&form.INPUT_TOTAL_TRANSATION=0&form.field1.earliest=-24h%40h&form.field1.latest=now&form.APPLICATION_NAME=CS2-ACZ-COSCON-PROD&form.MODULE_NAME=*&form.SERVICE_NAME=*&form.COMPARISON_TIME_RANGE=86400%23172800'
html = driver.get(url)
time.sleep(2)
##登录
username = driver.find_element_by_id('username')
username.clear()
username.send_keys(u'guowa')
time.sleep(1)
password = driver.find_element_by_id('password')
password.clear()
password.send_keys(u'123456')
time.sleep(1)
driver.find_element_by_xpath(u"html/body/div[2]/div/div/div[1]/form/fieldset/input[1]").submit()
time.sleep(23)
##把html 类名为menus的样式改为可见
js = 'document.getElementsByClassName("menus")[0].style.display="block";'
driver.execute_script(js)
time.sleep(1)
##获取导出
export=driver.find_element_by_xpath(u".//*[@id='element1-footer-8102']/div/a[2]/i")
time.sleep(1)
export.click()
##把样式改回来
js = 'document.getElementsByClassName("menus")[0].style.display="hide";'
driver.execute_script(js)
time.sleep(1)
##由于是弹窗，改变当前页面的句柄，把当前页面至首
for handle in driver.window_handles:
    driver.switch_to_window(handle)
##修改文件名和下载
driver.find_element_by_name("fileName").send_keys("CS2-ACZ-COSCON-PROD")
driver.find_element_by_class_name("modal-btn-primary").click()
time.sleep(3)

###get the time
dt = datetime.datetime.now()
current_year = int('20'+ dt.strftime('%y'))
current_mon = int(dt.strftime('%m'))
current_day = int(dt.strftime('%d')) - 1
begin_timelist = (current_year,current_mon,current_day,0,0,0,0,0,0)
last_timelist = (current_year,current_mon,current_day,24,0,0,0,0,0)
begin_time_str = str(int(time.mktime(begin_timelist)))
last_time_str = str(int(time.mktime(last_timelist)))

##获取CS2-ACZ-COSCONACZ-PROD的网址
url2 = 'http://artdc.cargosmart.com:8000/zh-CN/app/CS4_E2E_MON/CS2_Front_End_Process_Monitoring_ART_Report?form.INPUT_APPLICATION_NAME=CS2-ACZ-COSCONACZ-PROD&form.INPUT_MODULE_NAME=*&form.INPUT_SERVICE_NAME=*&form.INPUT_TOTAL_TRANSATION=0&form.field1.earliest='+begin_time_str+'&form.field1.latest='+last_time_str+'&form.APPLICATION_NAME=CS2-ACZ-COSCON-PROD&form.MODULE_NAME=*&form.SERVICE_NAME=*&form.COMPARISON_TIME_RANGE=86400%23172800'
html2 = driver.get(url2)
time.sleep(2)

##把html 类名为 menus的样式改为可见
js = 'document.getElementsByClassName("menus")[0].style.display="block";'
driver.execute_script(js)
time.sleep(10)
##点击导出
export=driver.find_element_by_xpath(u".//*[@id='element1-footer-8102']/div/a[2]/i")
time.sleep(2)
if export:
    export.click()
else:
    driver.quit()

##把样式改回来
js = 'document.getElementsByClassName("menus")[0].style.display="hide";'
driver.execute_script(js)
time.sleep(1)
##由于是弹窗，转换当前页面的句柄，把当前也至前
for handle in driver.window_handles:
    driver.switch_to_window(handle)
##输入文件名，点击下载
driver.find_element_by_name("fileName").send_keys("CS2-ACZ-COSCONACZ-PROD")
driver.find_element_by_class_name("modal-btn-primary").click()

time.sleep(4)
driver.quit()

