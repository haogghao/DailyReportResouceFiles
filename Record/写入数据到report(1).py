from urllib import request
from urllib import error
from urllib import parse
from http import cookiejar
import requests
from bs4 import BeautifulSoup
import time
import datetime
import selenium.webdriver.support.ui as ui
from selenium import webdriver
import sys
import xlrd
import xlwt
import time
import xdrlib,sys
import os
from xlutils.copy import copy
import win32com.client

def getTime():
    dt = datetime.datetime.now()
    Ydt = dt + datetime.timedelta(days=-1)
    currentYear = '20'+Ydt.strftime('%y')
    currentMon = Ydt.strftime('%m')
    currentDay = Ydt.strftime('%d')
    DailyFolderName = currentYear+currentMon+currentDay
    return DailyFolderName


def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception(e):
        print(str(e))


def getdaysRange():
    x = input("How many days to retrospect?(input 1 one day ago，input 2 two days ago，input 3 three days ago):")
    return x
        
def AutoGetColPosFator(daysRange):
    nowtime = datetime.datetime.now()
    needtime = nowtime + datetime.timedelta(days = 0-int(daysRange))
    #因为星期是0~6，所以为了通俗的显示，+1处理
    weekday = int(needtime.weekday())+1
    if (weekday+3) > 7:
        return weekday+3-7
    else:
        return weekday+3

    
def excel_get_zone_or_shipment(file,colnameindex,by_name):
    current_date = time.time() - (86400*1)
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

def excel_get_user(file,colnameindex = 0, by_name=u'Sheet0'):
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
    return list2[3:]



aczone_br_daily=0
aczone_br_edi=0
aczone_br_online=0
aczone_si_daily=0
aczone_si_edi=0
aczone_si_online=0

stdzone_br_daily=0
stdzone_br_edi=0
stdzone_br_online=0
stdzone_si_daily=0
stdzone_si_edi=0
stdzone_si_online=0

#进行total计算
def calc(list_aczone,list_stdzone):
    
    list_daily = []
    list_calc = []
    
    #声明全局变量
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


#写入excel数据
class WriteInExcel:  
      def __init__(self, filename=None):  #打开文件或者新建文件（如果不存在的话）
          self.xlApp = win32com.client.Dispatch('Excel.Application')  
          if filename:  
              self.filename = filename  
              self.xlBook = self.xlApp.Workbooks.Open(filename)  
          else:  
              self.xlBook = self.xlApp.Workbooks.Add()  
              self.filename = ''
      
      def save(self, newfilename=None):  #保存文件
          if newfilename:  
              self.filename = newfilename  
              self.xlBook.SaveAs(newfilename)  
          else:  
              self.xlBook.Save()      
      def close(self):  #关闭文件
          self.xlBook.Close(SaveChanges=0)  
          del self.xlApp

      def getCell(self, sheet, row, col):  #获取单元格的数据  
          "Get value of one cell"    
          sht = self.xlBook.Worksheets(sheet)    
          return sht.Cells(row, col).Value
        
      def setDateCell(self, sheet, row, col, value):  #设置单元格的数据  
          "set value of one cell"    
          sht = self.xlBook.Worksheets(sheet)    
          sht.Cells(row, col).Value = value  
                    
      '''
      lists:用数组作数据源
      multiplier:乘数，就是原数据每一行的列数
      row_in_sheet:在导入的表里开始位置的行坐标
      col_in_sheet:在导入的表里开始位置的列坐标
      sheet:创的表格
      '''
      def setCell(self,lists,multiplier,row_in_sheet,col_in_sheet,sheet):#设置单元格的数据
        sht = self.xlBook.Worksheets(sheet)
        begin = 0
        end = 0
        rows = len(lists)/multiplier
        for i in range(int(rows)):
            end = end + multiplier
            e_list = lists[begin:end]
            for j in range(len(e_list)):
                #sheet.write(row_in_sheet,j+col_in_sheet,e_list[j])
                d = sht.Cells(row_in_sheet, j+col_in_sheet)
                d.Value = (e_list[j])
            row_in_sheet = row_in_sheet + 1
            begin = begin + multiplier
            

      def clearRangeData(self, sheet, row1, col1, row2, col2):  #获得一块区域的数据，返回为一个二维元组,并清除为空  
          "return a 2d array (i.e. tuple of tuples)"    
          sht = self.xlBook.Worksheets(sheet)  
          sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value = ''
        

      #----------把网上数据写入excel表-----------
      def get_data(self,queueName,TYPE,start_total,end_total,sheet, row, col):#设置单元格的数据 
        #queueName：搜索数据的参数名称，如：Cargosmart.COSCON.CD.CT.INPUT.PARALLEL
        #TYPE：得出的某个数据的数量，如：CT的数量
        #start_total：开始的数据总量
        #end_total：结束的数据总量
        #---------CT，BC,BL,BL PRINT-----------
        wait.until(lambda driver:driver.find_element_by_id("from-time"))
        fromtime = driver.find_element_by_id("from-time")
        fromtime.clear()
        current_date = time.time() - (86400*1)
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
        driver.set_page_load_timeout(50)
        time.sleep(10) 
        html=driver.page_source#获取网页的html数据
        soup=BeautifulSoup(html,'html.parser')#对html进行解析，如果提示lxml未安装，直接pip install lxml即可
        table=soup.find("table",{'id':"emshist-tb-1"})
        flag=0#标记，当爬取字段数据是为0，否则为1
        num=len(table.find_all('tr'))
        print(TYPE)
        for tr in table.find_all('tr'): 
            if flag==1:     #第一行为表格字段数据，因此跳过第一行
                dic={}
                flag+=1
                start_total = tr.find_all('td')[5].string
                print(start_total)
            if flag==num:
                end_total = tr.find_all('td')[5].string
                print(end_total)
                flag+=1
            flag+=1
        driver.set_page_load_timeout(50)
        time.sleep(2) 
        "set value of one cell"    
        sht = self.xlBook.Worksheets(sheet)
        total = int(end_total) - int(start_total)
        print(total)
        sht.Cells(row,col).Value = total
        
if __name__ == "__main__":
    daysRange = getdaysRange()
    weekday_num = AutoGetColPosFator(daysRange)
    today=time.strftime("%m%d",time.localtime(time.time()))
    DailyFolderName = getTime()
    xls = WriteInExcel(r"D:\DailyReportResouceFiles\Report\COSCON-ACZ-daily-stat-result "+today+".xlsx")
    UserSync = WriteInExcel(r'D:\\DailyReportResouceFiles\\'+DailyFolderName+'\\Report - Coscon User Profile Sync Txn Report.xlsx')
    #写入aczong的 #写入stdzone的
    list_aczone = excel_get_zone_or_shipment('D:\\DailyReportResouceFiles\\'+DailyFolderName+'\\ACZone TXN Monitor.xlsx',0,u'Sheet0')
    list_stdzone = excel_get_zone_or_shipment('D:\\DailyReportResouceFiles\\'+DailyFolderName+'\\STDZone COSCON BR SI Daily TXN Report.xlsx',0,u'Sheet0')
    list_daily_result = calc(list_aczone,list_stdzone)[0]
    list_calc_result = calc(list_aczone,list_stdzone)[1]
    xls.setCell(list_daily_result,1,19,int(weekday_num*2+2),'Shipments')
    xls.setCell(list_calc_result,1,8,int(weekday_num+1),'Shipments')
    

    #写入shipment 
    list_shipment = excel_get_zone_or_shipment('D:\\DailyReportResouceFiles\\'+DailyFolderName+'\\ACZone Shipment Folder Txn Report.xlsx',0,u'Sheet0')
    #如果是判断周五的话，就进行日期的更新
    if  weekday_num == 1:
        xls.clearRangeData('Shipments', 3, 2, 12, 8)
        xls.clearRangeData('Shipments', 19, 4, 27, 17)
        i = 2
        j = 2
        for j in range(2,9):
            if j < 5:
                xls.setDateCell('Shipments',2,j,time.strftime("%m/%d/%Y",time.localtime(time.time()-(86400*(5-j)))))
                xls.setDateCell('Shipments',17,i*j,time.strftime("%m/%d/%Y",time.localtime(time.time()-(86400*(5-j)))))
            else:
                xls.setDateCell('Shipments',2,j,time.strftime("%m/%d/%Y",time.localtime(time.time()+(86400*(j-5)))))
                xls.setDateCell('Shipments',17,i*j,time.strftime("%m/%d/%Y",time.localtime(time.time()+(86400*(j-5)))))
    e_list_shipment = [0]   
    if list_shipment:
        e_list_shipment[0] = list_shipment[1]
    else:
        e_list_shipment[0] = 0  
    xls.setCell(e_list_shipment,1,23,int(weekday_num*2+2),'Shipments')

    #写入user的
    list_user = excel_get_user('D:\\DailyReportResouceFiles\\'+DailyFolderName+'\\Report - Coscon User Profile Sync Txn Report.xlsx',0,u'Sheet0')
    #由于周一不能清空数据，所以要插入数据之前必须先进行比较
##    for i in range(32,42):
##        if xls.getCell('UserSync',i,1) == UserSync.getCell('Sheet0',i-30,1) and xls.getCell('UserSync',i,2) == UserSync.getCell('Sheet0',i-30,2):
##           xls.setDateCell('UserSync',i,3,UserSync.getCell('Sheet0',i-30,3))
##        elif xls.getCell('UserSync',i,1) == UserSync.getCell('Sheet0',i-30,1) and xls.getCell('UserSync',i,2) != UserSync.getCell('Sheet0',i-30,2):
           #xls.setDateCell('UserSync',i,3,0)
    #如果是判断周五的话，就进行日期的更新
    if weekday_num == 1:
       xls.clearRangeData('UserSync', 5, 7, 16, 13)
       xls.clearRangeData('UserSync', 19, 1, 29, 13)
       xls.clearRangeData('UserSync', 32, 1, 45, 3)
       #更新时间
       xls.setDateCell('UserSync',3,1,time.strftime("%m/%d/%Y",time.localtime(time.time()-(86400*3))))
       xls.setDateCell('UserSync',3,7,time.strftime("%m/%d/%Y",time.localtime(time.time()-(86400*2))))
       xls.setDateCell('UserSync',3,11,time.strftime("%m/%d/%Y",time.localtime(time.time()-(86400*1))))
       xls.setDateCell('UserSync',17,1,time.strftime("%m/%d/%Y",time.localtime(time.time()+(86400*0))))
       xls.setDateCell('UserSync',17,7,time.strftime("%m/%d/%Y",time.localtime(time.time()+(86400*1))))
       xls.setDateCell('UserSync',17,11,time.strftime("%m/%d/%Y",time.localtime(time.time()+(86400*2))))
       xls.setDateCell('UserSync',30,1,time.strftime("%m/%d/%Y",time.localtime(time.time()+(86400*3))))

    if weekday_num < 4: 
       write_in_row = 5
       write_in_column = weekday_num*4-1
    elif weekday_num >= 4 and weekday_num < 7:
        if weekday_num == 4:
           write_in_row = 19
           write_in_column = 4*(weekday_num-3)-3
        else:
           write_in_row = 19
           write_in_column = 4*(weekday_num-3)-1
    else:
        write_in_row = 32
        write_in_column = 1
    xls.setCell(list_user,3,write_in_row,write_in_column,'UserSync')


    

##    #爬取网上的数据
##    #登录地址
##    login_url = 'https://www.cargosmart.com/operation/j_spring_security_check'
##    get_url = 'https://www.cargosmart.com/operation/ems/ems_history.jsf'
##    driver = webdriver.Firefox()
##    driver.maximize_window()
##    wait = ui.WebDriverWait(driver,10)
##    #driver.set_window_size(1055, 800)  
##    driver.get(login_url)
##    username = driver.find_element_by_name("j_username")
##    username.clear()
##    username.send_keys('ocuser')
##    password = driver.find_element_by_name("j_password")
##    password.clear()
##    password.send_keys('ocuser')
##    time.sleep(1)
##    driver.find_element_by_xpath("//*[@id='submit']").click() 
##    time.sleep(5)
##    html=driver.get(get_url)
##    html=driver.page_source#获取网页的html数据
##    soup=BeautifulSoup(html,"html.parser")#对html进行解析，如果提示lxml未安装，直接pip install lxml即可
##    xls.get_data("Cargosmart.COSCON.CD.CT.INPUT.PARALLEL","CT的数量","CT_start_total","CT_end_total","Shipments",24,int(weekday_num*2+2))
##    time.sleep(5)
##    xls.get_data("Cargosmart.COSCON.CD.CT.INPUT.PARALLEL","CT的数量","CT_start_total","CT_end_total","Shipments",3,int(weekday_num+1))
##    time.sleep(5)
##    xls.get_data("Cargosmart.COSCON.CD.BC.INPUT.PARALLEL","BC的数量","BC_start_total","BC_end_total","Shipments",25,int(weekday_num*2+2))
##    time.sleep(5)
##    xls.get_data("Cargosmart.COSCON.CD.BC.INPUT.PARALLEL","BC的数量","BC_start_total","BC_end_total","Shipments",4,int(weekday_num+1))
##    time.sleep(5)
##    xls.get_data("Cargosmart.COSCON.CD.BL.INPUT.PARALLEL","BL的数量","BL_start_total","BL_end_total","Shipments",26,int(weekday_num*2+2))
##    time.sleep(5)
##    xls.get_data("Cargosmart.COSCON.CD.BL.INPUT.PARALLEL","BL的数量","BL_start_total","BL_end_total","Shipments",5,int(weekday_num+1))
##    time.sleep(5)
##    xls.get_data("CS2.BILLOFLADING.PRINT.COSU.IN.QUE","BL Print的数量","BL_PRINT_start_total","BL_PRINT_end_total","Shipments",27,int(weekday_num*2+2))
##    time.sleep(5)
##    xls.get_data("CS2.BILLOFLADING.PRINT.COSU.IN.QUE","BL Print的数量","BL_PRINT_start_total","BL_PRINT_end_total","Shipments",6,int(weekday_num+1))
##    time.sleep(5)
##    driver.close()  
##    driver.quit()
  
    xls.save()
    xls.close()
    
