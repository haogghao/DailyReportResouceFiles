#encoding=utf-8
import xlrd
import time
import xdrlib,sys
import xlwt
from bs4 import BeautifulSoup#bill add
from selenium import webdriver
import selenium.webdriver.support.ui as ui
from selenium.webdriver.common.keys import Keys


def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception(e):
        print(str(e))
           
def excel_get_zone_or_shipment(file,colnameindex,by_name,x):
    current_date = 0
    if(int(x) == 1):
        current_date = time.time() - 86400
    elif (int(x) == 2):
        current_date = time.time() - (86400*int(x))
    elif(int(x) == 3):
        current_date = time.time() - (86400*int(x))
    #current_date = time.time() - 86400
    timeStr = time.strftime("%y-%m-%d",time.localtime(current_date))
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows
    colnames = table.row_values(colnameindex)
    list = []
    list2 = []
    for rownum in range(0,nrows):
        row = table.row_values(rownum)
        if row:
            if (str(table.cell(rownum,0)) == "text:'20"+str(timeStr)+"'"):
                app = {}
                for i in range(len(colnames)):
                    app[colnames[i]] = row[i]
                list.append(app)
    
    for i in range(len(list)):
            for j in range(len(list[i])):
                    list2.append(list[i][colnames[j]])
    return list2

def excel_get_user(file = 'Report - Coscon User Profile Sync Txn Report.xlsx',colnameindex = 0, by_name=u'Sheet0'):
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows
    colnames = table.row_values(colnameindex)
    list = []
    list2 = []
    for rownum in range(0,nrows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            list.append(app)
    for i in range(len(list)):
        for j in range(len(list[i])):
            list2.append(list[i][colnames[j]])
    return list2

'''
lists:用数组作数据源
multiplier:乘数，就是原数据每一行的列数
row_in_sheet:在导入的表里开始位置的行坐标
col_in_sheet:在导入的表里开始位置的列坐标
sheet:创的表格
'''
def write_in_excel(lists,multiplier,row_in_sheet,col_in_sheet,sheet):
    begin = 0
    end = 0
    #var = get_rows(lists,multiplier)
    rows = len(lists)/multiplier
    for i in range(int(rows)):
        end = end + multiplier
        e_list = lists[begin:end]
        for j in range(len(e_list)):
            sheet.write(row_in_sheet,j+col_in_sheet,e_list[j])
        row_in_sheet = row_in_sheet + 1
        begin = begin + multiplier

#全局变量
aczone_br_daily = 0
aczone_br_edi = 0
aczone_br_online = 0
aczone_si_daily = 0
aczone_si_edi = 0
aczone_si_online = 0

stdzone_br_daily = 0
stdzone_br_edi = 0
stdzone_br_online = 0
stdzone_si_daily = 0
stdzone_si_edi = 0
stdzone_si_online = 0

def calc(list_aczone,list_stdzone):
    
    list_daily = []
    list_calc = []
    
    #声明
    global aczone_br_daily
    global aczone_br_edi
    global aczone_br_online
    global aczone_si_daily
    global aczone_si_edi
    global aczone_si_online

    global stdzone_br_daily
    global stdzone_br_edi
    global stdzone_br_online
    global stdzone_si_daily
    global stdzone_si_edi
    global stdzone_si_online

    for i in range(len(list_aczone)):
        if list_aczone[i] == 'BR':
            aczone_br_daily = list_aczone[i-1]
            aczone_br_edi = list_aczone[i+1]
            aczone_br_online = list_aczone[i+2]
            
        if list_aczone[i] == 'SI':
            aczone_si_daily = list_aczone[i-1]
            aczone_si_edi = list_aczone[i+1]
            aczone_si_online = list_aczone[i+2]

    for j in range(len(list_stdzone)):
        if list_stdzone[j] == 'BR':
            stdzone_br_daily = list_stdzone[j-1]
            stdzone_br_edi = list_stdzone[j+1]
            stdzone_br_online = list_stdzone[j+2]
            
        if list_stdzone[j] == 'SI':
            stdzone_si_daily = list_stdzone[j-1]
            stdzone_si_edi = list_stdzone[j+1]
            stdzone_si_online = list_stdzone[j+2]
    
    list_daily.append(stdzone_br_daily)
    list_daily.append(aczone_br_daily)
    list_daily.append(stdzone_si_daily)
    list_daily.append(aczone_si_daily)

    list_calc.append(int(aczone_br_edi)+int(stdzone_br_edi))
    list_calc.append(int(aczone_br_online)+int(stdzone_br_online))
    list_calc.append(int(aczone_si_online)+int(stdzone_si_online))
    list_calc.append(int(aczone_si_edi)+int(stdzone_si_edi))
    list_calc.append(int(aczone_br_edi)+int(stdzone_br_edi)+int(aczone_br_online)+int(stdzone_br_online)+int(aczone_si_online)+int(stdzone_si_online)+int(aczone_si_edi)+int(stdzone_si_edi))

    result = [list_daily,list_calc]
    
    return result

def get_CT_BC_BL_BLPrint(url,Dtime):
    CT_BC_BL_BLPrint=[]
    driver=webdriver.Firefox()
    wait = ui.WebDriverWait(driver,10)
    driver.set_page_load_timeout(30)
    time.sleep(3)
    html=driver.get(url[0])#使用get方法请求url，因为是模拟浏览器，所以不需要headers信息
    wait.until(lambda driver: driver.find_element_by_name(u"j_username"))
    username=driver.find_element_by_name(u"j_username")#登录
    username.clear()
    username.send_keys(u'ocuser')
    wait.until(lambda driver: driver.find_element_by_name(u"j_password"))
    pd=driver.find_element_by_name(u"j_password")
    pd.clear()
    pd.send_keys(u'ocuser')
    wait.until(lambda driver: driver.find_element_by_id(u"submit"))
    driver.find_element_by_id(u"submit").click()#提交登录
    driver.set_page_load_timeout(30)
    time.sleep(5)
    html=driver.get(url[1])
    wait.until(lambda driver:driver.find_element_by_id(u"from-time"))
    fromtime = driver.find_element_by_id(u"from-time")
    fromtime.clear()
    fromtime.send_keys(Dtime[0])
    time.sleep(1)
    wait.until(lambda driver:driver.find_element_by_id(u"to-time"))
    totime = driver.find_element_by_id(u"to-time")
    totime.clear()
    totime.send_keys(Dtime[1])
    time.sleep(1)
    wait.until(lambda driver:driver.find_element_by_name(u"emsInstance"))
    emsInstance = driver.find_element_by_name(u"emsInstance")
    emsInstance.clear()
    emsInstance.send_keys(u'tcp://cosems01v:8001')
    time.sleep(1)
#######################----------获取 CT------------#######
    queueKey=u'Cargosmart.COSCON.CD.CT.INPUT.PARALLEL'
    CT=get_first_and_last_data(queueKey,driver)
    CT_BC_BL_BLPrint.append(CT)
######################  获取BC  ###################
    queueKey=u'Cargosmart.COSCON.CD.BC.INPUT.PARALLEL'
    BC=get_first_and_last_data(queueKey,driver)
    CT_BC_BL_BLPrint.append(BC)
###################  获取BL  #############
    queueKey=u'Cargosmart.COSCON.CD.BL.INPUT.PARALLEL'
    BL=get_first_and_last_data(queueKey,driver)
    CT_BC_BL_BLPrint.append(BL)
####################  获取BLPrint  ###########
    queueKey=u'CS2.BILLOFLADING.PRINT.COSU.IN.QUE'
    BLPrint=get_first_and_last_data(queueKey,driver)
    CT_BC_BL_BLPrint.append(BLPrint)
    return CT_BC_BL_BLPrint


def get_first_and_last_data(queueKey,driver):
    wait = ui.WebDriverWait(driver,10)
    wait.until(lambda driver:driver.find_element_by_name(u"queueName"))
    queuename = driver.find_element_by_name(u"queueName")
    queuename.clear()
    queuename.send_keys(queueKey)
    time.sleep(1)
    wait.until(lambda driver:driver.find_element_by_xpath(u"//*[@id='history-search-criteria']/form/table/tbody/tr[4]/td/input[2]"))
    driver.find_element_by_xpath(u"//*[@id='history-search-criteria']/form/table/tbody/tr[4]/td/input[2]").send_keys(Keys.ENTER)
    time.sleep(20)
    wait.until(lambda driver:driver.page_source)
    html=driver.page_source#获取网页的html数据，自动更新，只需要重新获取driver.page_source即可。
    soup=BeautifulSoup(html,"html.parser")#对html进行解析，如果提示lxml未安装，直接pip install lxml即可
    table=soup.find("table",{'id':"emshist-tb-1"})
    i=0#标记，0表示第一行，n表示最后一行
    num=len(table.find_all('tr'))
    print('总共 %s 行数据' % (num-1))
    data_list=[]
    for tr in table.find_all('tr'): 
        if i==1:     #第一行为表格字段数据，因此跳过第一行
            dic={}
            i+=1
            print(tr.find_all('td')[5].string)
            data_list.append(int(tr.find_all('td')[5].string))
        if i==num:
            print(tr.find_all('td')[5].string)
            data_list.append(int(tr.find_all('td')[5].string))
            i+=1
        i+=1
    return data_list[1]-data_list[0]



    
def main():    
    wbk_1 = xlwt.Workbook()
    sheet_1 = wbk_1.add_sheet('new',cell_overwrite_ok = True)
    x = input("How many days ago would you like to check the data：（input 1 one day ago，input 2 two days ago，input 3 three days ago")
    
    #写入aczong的
    list_aczone = excel_get_zone_or_shipment('ACZone TXN Monitor.xlsx',0,u'Sheet0',x)
    sheet_1.write(0,0,'aczone')
    sheet_1.write(1,0,'date')
    sheet_1.write(1,1,'daily')
    sheet_1.write(1,2,'type')
    sheet_1.write(1,3,'edi')
    sheet_1.write(1,4,'online')
    write_in_excel(list_aczone,5,2,0,sheet_1)

    #写入shipment
    list_shipment = excel_get_zone_or_shipment('ACZone Shipment Folder Txn Report.xlsx',0,u'Sheet0',x)
    sheet_1.write(5,0,'shipment')
    sheet_1.write(6,0,'date')
    sheet_1.write(6,1,'count')
    write_in_excel(list_shipment,2,7,0,sheet_1)
    
    #写入stdzone的
    list_stdzone = excel_get_zone_or_shipment('STDZone COSCON BR SI Daily TXN Report.xlsx',0,u'Sheet0',x)
    sheet_1.write(0,7,'stdzone')
    sheet_1.write(1,7,'date')
    sheet_1.write(1,8,'daily')
    sheet_1.write(1,9,'type')
    sheet_1.write(1,10,'edi')
    sheet_1.write(1,11,'online')
    write_in_excel(list_stdzone,5,2,7,sheet_1)

    #写入user的
    list_user = excel_get_user()
    sheet_1.write(5,3,'user')
    write_in_excel(list_user,3,6,3,sheet_1)

    #写入需要用到的值
    sheet_1.write(20,0,'std_br_daily')
    sheet_1.write(21,0,'acz_br_daily')
    sheet_1.write(22,0,'std_si_daily')
    sheet_1.write(23,0,'acz_si_daily')
    
    sheet_1.write(20,9,'edi_br')
    sheet_1.write(21,9,'online_br')
    sheet_1.write(22,9,'online_si')
    sheet_1.write(23,9,'edi_si')
    sheet_1.write(24,9,'total')

    list_daily_result = calc(list_aczone,list_stdzone)[0]
    list_calc_result = calc(list_aczone,list_stdzone)[1]

    write_in_excel(list_daily_result,1,20,2,sheet_1)
    write_in_excel(list_calc_result,1,20,11,sheet_1)

    #写入CT,BC,BL,BLPrint
    url=['https://www.cargosmart.com/operation/index.jsf','https://www.cargosmart.com/operation/ems/ems_history.jsf']
    current_date = time.time()
    from_time = time.strftime("%d %b %Y 00:00",time.localtime(current_date - 86400*int(x))) #格式为16 Aug 2017 00:00
    to_time= time.strftime("%d %b %Y 00:05",time.localtime(current_date-86400*(int(x)-1)))
    Dtime=[from_time,to_time]
    
    
    sheet_1.write(28,0,'CT')
    sheet_1.write(29,0,'BC')
    sheet_1.write(30,0,'BL')
    sheet_1.write(31,0,'BLPrint')

    resultList=get_CT_BC_BL_BLPrint(url,Dtime)#调用函数
    write_in_excel(resultList,1,28,1,sheet_1)
    
    sheet_1.write(28,9,resultList[0])
    sheet_1.write(29,9,resultList[1])
    sheet_1.write(30,9,resultList[2])
    sheet_1.write(31,9,resultList[3])
    
    wbk_1.save('hi.xls')

    
main()

