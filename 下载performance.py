from urllib import request
import time
from urllib import error
from urllib import parse
from http import cookiejar
import requests
import re
import datetime
from selenium import webdriver
import selenium.webdriver.support.ui as ui
from bs4 import BeautifulSoup

def getTime():
    dt = datetime.datetime.now()
    Ydt = dt + datetime.timedelta(days=-1)
    currentYear = '20'+Ydt.strftime('%y')
    currentMon = Ydt.strftime('%m')
    currentDay = Ydt.strftime('%d')
    DailyFolderName = currentYear+currentMon+currentDay
    return DailyFolderName

def get():
    login_url="http://artdc.cargosmart.com:8000/zh-CN/app/CS4_E2E_MON/CS2_Front_End_Process_Monitoring_ART_Report?form.INPUT_APPLICATION_NAME=CS2-ACZ-COSCON-PROD&form.INPUT_MODULE_NAME=*&form.INPUT_SERVICE_NAME=*&form.INPUT_TOTAL_TRANSATION=0&form.field1.earliest=-24h%40h&form.field1.latest=now&form.APPLICATION_NAME=CS2-ACZ-COSCON-PROD&form.MODULE_NAME=*&form.SERVICE_NAME=*&form.COMPARISON_TIME_RANGE=86400%23172800"
    login_url_coscon = "http://artdc.cargosmart.com:8000/zh-CN/app/CS4_E2E_MON/search?earliest=-24h%40h&latest=now&sid=guowa__guowa_Q1M0X0UyRV9NT04__search4_1503566086.43355&q=search%20index%3D%22artraw%22%20sourcetype%3D%22artraw-csv%22%7C%20where%20!isnull(ELAPSED_TIME)%20%7C%20where%20APPLICATION_NAME%20%3D%20if(%22CS2-ACZ-COSCON-PROD%22%3D%3D%22*%22%2CAPPLICATION_NAME%2C%22CS2-ACZ-COSCON-PROD%22)%20%7C%20where%20MODULE_NAME%20%3D%20if(%22*%22%3D%3D%22*%22%2CMODULE_NAME%2C%22*%22)%20%7C%20where%20SERVICE_NAME%20%3D%20if(%22*%22%3D%3D%22*%22%2CSERVICE_NAME%2C%22*%22)%20%7C%20%20stats%20count%20as%20%22TOTAL_TRANSATION%22%2C%20avg(ELAPSED_TIME)%20as%20%22AVERAGE_RESPONSE_TIME%22%2C%20count(eval(if(ELAPSED_TIME%20%3C%3D%205%2C1%2CNULL)))%20AS%20%22LESS_THAN_5s%22%2C%20count(eval(if(ELAPSED_TIME%20%3C%3D%2010%20AND%20ELAPSED_TIME%20%3E%3D%205%2C1%2CNULL)))%20AS%20%22BETWEEN_5s_TO_10s%22%2C%20count(eval(if(ELAPSED_TIME%20%3C%3D%2015%20AND%20ELAPSED_TIME%20%3E%3D%2010%2C1%2CNULL)))%20AS%20%22BETWEEN_10s_TO_15s%22%2C%20count(eval(if(ELAPSED_TIME%20%3C%3D%2020%20AND%20ELAPSED_TIME%20%3E%3D%2015%2C1%2CNULL)))%20AS%20%22BETWEEN_15s_TO_20s%22%2Ccount(eval(if(ELAPSED_TIME%20%3E%3D%2020%2C1%2CNULL)))%20AS%20%22MORE_THAN_20s%22%20by%20APPLICATION_NAME%2CMODULE_NAME%2CSERVICE_NAME%20%7C%20where%20TOTAL_TRANSATION%20%3E%3D%200%20%7C%20eval%20RATE_LESS_THAN_5s%3Dsubstr(tostring(LESS_THAN_5s%2FTOTAL_TRANSATION*100)%2C1%2C4)%2B%22%25%22%20%7C%20eval%20RATE_BETWEEN_5s_TO_10s%3Dsubstr(tostring(BETWEEN_5s_TO_10s%2FTOTAL_TRANSATION*100)%2C1%2C4)%2B%22%25%22%20%7C%20%20eval%20RATE_BETWEEN_10s_TO_15s%3Dsubstr(tostring(BETWEEN_10s_TO_15s%2FTOTAL_TRANSATION*100)%2C1%2C4)%2B%22%25%22%20%7C%20%20eval%20RATE_BETWEEN_15s_TO_20s%3Dsubstr(tostring(BETWEEN_15s_TO_20s%2FTOTAL_TRANSATION*100)%2C1%2C4)%2B%22%25%22%20%7C%20eval%20RATE_MORE_THAN_20s%3Dsubstr(tostring(MORE_THAN_20s%2FTOTAL_TRANSATION*100)%2C1%2C4)%2B%22%25%22%20%20%7C%20table%20APPLICATION_NAME%2CMODULE_NAME%2CSERVICE_NAME%2CTOTAL_TRANSATION%2CAVERAGE_RESPONSE_TIME%2CLESS_THAN_5s%2CRATE_LESS_THAN_5s%2CBETWEEN_5s_TO_10s%2CRATE_BETWEEN_5s_TO_10s%2CBETWEEN_10s_TO_15s%2CRATE_BETWEEN_10s_TO_15s%2CBETWEEN_15s_TO_20s%2CRATE_BETWEEN_15s_TO_20s%2CMORE_THAN_20s%2CRATE_MORE_THAN_20s&display.general.type=statistics&display.page.search.mode=fast&dispatch.sample_ratio=1"
    login_url_cosconacz="http://artdc.cargosmart.com:8000/zh-CN/app/CS4_E2E_MON/search?earliest=-1d%40d&latest=%40d&sid=guowa__guowa_Q1M0X0UyRV9NT04__search4_1503642503.48202&q=search%20index%3D%22artraw%22%20sourcetype%3D%22artraw-csv%22%7C%20where%20!isnull(ELAPSED_TIME)%20%7C%20where%20APPLICATION_NAME%20%3D%20if(%22CS2-ACZ-COSCONACZ-PROD%22%3D%3D%22*%22%2CAPPLICATION_NAME%2C%22CS2-ACZ-COSCONACZ-PROD%22)%20%7C%20where%20MODULE_NAME%20%3D%20if(%22*%22%3D%3D%22*%22%2CMODULE_NAME%2C%22*%22)%20%7C%20where%20SERVICE_NAME%20%3D%20if(%22*%22%3D%3D%22*%22%2CSERVICE_NAME%2C%22*%22)%20%7C%20%20stats%20count%20as%20%22TOTAL_TRANSATION%22%2C%20avg(ELAPSED_TIME)%20as%20%22AVERAGE_RESPONSE_TIME%22%2C%20count(eval(if(ELAPSED_TIME%20%3C%3D%205%2C1%2CNULL)))%20AS%20%22LESS_THAN_5s%22%2C%20count(eval(if(ELAPSED_TIME%20%3C%3D%2010%20AND%20ELAPSED_TIME%20%3E%3D%205%2C1%2CNULL)))%20AS%20%22BETWEEN_5s_TO_10s%22%2C%20count(eval(if(ELAPSED_TIME%20%3C%3D%2015%20AND%20ELAPSED_TIME%20%3E%3D%2010%2C1%2CNULL)))%20AS%20%22BETWEEN_10s_TO_15s%22%2C%20count(eval(if(ELAPSED_TIME%20%3C%3D%2020%20AND%20ELAPSED_TIME%20%3E%3D%2015%2C1%2CNULL)))%20AS%20%22BETWEEN_15s_TO_20s%22%2Ccount(eval(if(ELAPSED_TIME%20%3E%3D%2020%2C1%2CNULL)))%20AS%20%22MORE_THAN_20s%22%20by%20APPLICATION_NAME%2CMODULE_NAME%2CSERVICE_NAME%20%7C%20where%20TOTAL_TRANSATION%20%3E%3D%200%20%7C%20eval%20RATE_LESS_THAN_5s%3Dsubstr(tostring(LESS_THAN_5s%2FTOTAL_TRANSATION*100)%2C1%2C4)%2B%22%25%22%20%7C%20eval%20RATE_BETWEEN_5s_TO_10s%3Dsubstr(tostring(BETWEEN_5s_TO_10s%2FTOTAL_TRANSATION*100)%2C1%2C4)%2B%22%25%22%20%7C%20%20eval%20RATE_BETWEEN_10s_TO_15s%3Dsubstr(tostring(BETWEEN_10s_TO_15s%2FTOTAL_TRANSATION*100)%2C1%2C4)%2B%22%25%22%20%7C%20%20eval%20RATE_BETWEEN_15s_TO_20s%3Dsubstr(tostring(BETWEEN_15s_TO_20s%2FTOTAL_TRANSATION*100)%2C1%2C4)%2B%22%25%22%20%7C%20eval%20RATE_MORE_THAN_20s%3Dsubstr(tostring(MORE_THAN_20s%2FTOTAL_TRANSATION*100)%2C1%2C4)%2B%22%25%22%20%20%7C%20table%20APPLICATION_NAME%2CMODULE_NAME%2CSERVICE_NAME%2CTOTAL_TRANSATION%2CAVERAGE_RESPONSE_TIME%2CLESS_THAN_5s%2CRATE_LESS_THAN_5s%2CBETWEEN_5s_TO_10s%2CRATE_BETWEEN_5s_TO_10s%2CBETWEEN_10s_TO_15s%2CRATE_BETWEEN_10s_TO_15s%2CBETWEEN_15s_TO_20s%2CRATE_BETWEEN_15s_TO_20s%2CMORE_THAN_20s%2CRATE_MORE_THAN_20s&display.page.search.mode=fast&dispatch.sample_ratio=1&display.general.type=statistics"
    user_agent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:52.0) Gecko/20100101 Firefox/52.0"
    #Headers信息
    head = {'User-Agent': user_agent,'Connection': "keep-alive"}
    #登陆Form_Data信息
    Login_Data = {}
    Login_Data['username'] = 'guowa'
    Login_Data['password'] = '123456'
    #使用urlencode方法转换标准格式
    logingpostdata = parse.urlencode(Login_Data).encode('utf-8')
    #声明一个CookieJar对象实例来保存cookie
    cookie_filename = 'cookie_jar.txt'
    cookie = cookiejar.MozillaCookieJar(cookie_filename)
    #利用urllib.request库的HTTPCookieProcessor对象来创建cookie处理器,也就CookieHandler
    cookie_support = request.HTTPCookieProcessor(cookie)
    #通过CookieHandler创建opener
    opener = request.build_opener(cookie_support)
    #创建Request对象
    req = request.Request(url=login_url, data=logingpostdata, headers=head)
    try:
        response = opener.open(req)
        html = response.read().decode('utf-8')
        #print(html)
    except error.URLError as e:
        if hasattr(e, 'code'):
            print("HTTPError:%d" % e.code)
        elif hasattr(e, 'reason'):
            print("URLError:%s" % e.reason)
    DailyFolderName = getTime()
    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups':0,'download.default_directory':'D:\\DailyReportResouceFiles\\'+DailyFolderName}
    options.add_experimental_option('prefs',prefs)
    driver = webdriver.Chrome(chrome_options=options)
    driver.maximize_window()
    wait = ui.WebDriverWait(driver,10)
    #下载CS2-ACZ-COSCON-PROD表格
    driver.get(login_url_coscon)
    time.sleep(1)
    wait.until(lambda driver:driver.find_element_by_id("username"))
    username = driver.find_element_by_id("username")
    username.clear()
    username.send_keys('guowa')
    time.sleep(1)
    wait.until(lambda driver:driver.find_element_by_id("password"))
    password = driver.find_element_by_id("password")
    password.clear()
    password.send_keys('123456')
    time.sleep(1)
    driver.find_element_by_xpath("html/body/div[2]/div/div/div[1]/form/fieldset/input[1]").click()
    driver.set_page_load_timeout(50)
    time.sleep(20)
    driver.find_element_by_xpath("html/body/div[1]/div/div/div[1]/div[4]/div/div[1]/div[2]/div[2]/a[3]/i").click()
    driver.set_page_load_timeout(50)
    time.sleep(5)
    wait.until(lambda driver:driver.find_element_by_name("fileName"))
    fileName = driver.find_element_by_name("fileName")
    fileName.clear()
    fileName.send_keys("CS2-ACZ-COSCON-PROD.csv")
    time.sleep(1)
    driver.find_element_by_xpath("html/body/div[4]/div[3]/a[2]").click()
    driver.set_page_load_timeout(50)
    time.sleep(1)
    html=driver.page_source#获取网页的html数据
    soup=BeautifulSoup(html,'html.parser')#对html进行解析，如果提示lxml未安装，直接pip install lxml即可


    
    #下载CS2-ACZ-COSCONACZ-PROD表格
    #登陆主界面
    driver.get(login_url)
    driver.set_page_load_timeout(50)
    time.sleep(10)
    #搜索下拉框
    wait.until(lambda driver:driver.find_element_by_id("select2-chosen-2"))
    driver.find_element_by_id("select2-chosen-2").click()
    time.sleep(1)
    wait.until(lambda driver:driver.find_element_by_id("s2id_autogen2_search"))
    drop_down_input = driver.find_element_by_id("s2id_autogen2_search")
    drop_down_input.clear()
    drop_down_input.send_keys('CS2-ACZ-COSCONACZ-PROD')
    time.sleep(1)
    driver.find_element_by_id("select2-results-2").click()
    driver.set_page_load_timeout(50)
    time.sleep(5)
    driver.find_element_by_class_name("time-label").click()
    driver.set_page_load_timeout(50)
    time.sleep(5)
    driver.find_element_by_link_text("昨天").click()
    driver.set_page_load_timeout(50)
    time.sleep(5)


    #进入下载界面
    driver.get(login_url_cosconacz)
    driver.set_page_load_timeout(50)
    time.sleep(20)
    driver.find_element_by_xpath("html/body/div[1]/div/div/div[1]/div[4]/div/div[1]/div[2]/div[2]/a[3]/i").click()
    driver.set_page_load_timeout(50)
    time.sleep(3)
    wait.until(lambda driver:driver.find_element_by_name("fileName"))
    fileName = driver.find_element_by_name("fileName")
    fileName.clear()
    fileName.send_keys("CS2-ACZ-COSCONACZ-PROD.csv")
    time.sleep(1)
    driver.find_element_by_xpath("html/body/div[4]/div[3]/a[2]").click()
    driver.set_page_load_timeout(50)
    time.sleep(5)
    html=driver.page_source#获取网页的html数据
    soup=BeautifulSoup(html,'html.parser')#对html进行解析，如果提示lxml未安装，直接pip install lxml即可
    driver.close()  
    driver.quit()
    
get()






    
