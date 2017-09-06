#!/usr/bin/env python   
# -*- coding: utf-8 -*- 
from win32com.client import Dispatch
import win32com.client
import time
import os
import openpyxl
import sys

yesterday= time.strftime("%Y%m%d",time.localtime(time.time()-86400))
today=time.strftime("%m%d",time.localtime(time.time()))
ytd=time.strftime("%a",time.localtime(time.time()-86400))
pic1=r'\\'.join(['D:','DailyReportResouceFiles',yesterday,'COSCON Network Utilization','5min.png'])
pic2=r'\\'.join(['D:','DailyReportResouceFiles',yesterday,'COSCON Network Utilization','30min.png'])
excelPath=r'\\'.join(['D:','DailyReportResouceFiles','Report','COSCON-ACZ-daily-stat-result '+today+'.xlsx'])
#excelPath2="D:\DailyReportResouceFiles\Report\COSCON-ACZ-daily-stat-result "+today+".xlsx"
insert_date=time.strftime("%Y/%b/%d",time.localtime(time.time()-86400))+' COSCON 10M lease line usage : < 25%'
      
def Mon(xls):
      xls.addPicture('Network', pic1, 0,500,389,170)
      xls.addPicture('Network', pic2, 0,735,389,170)
      xls.setCell(32,1,insert_date)
      xls.setCell(47,3,'Daily (5 minutes average)')
      xls.setCell(62,3,'Weekly (30 minutes average)')
      save_close(xls)
def Tue(xls):
      xls.addPicture('Network', pic1, 480,500,389,170)
      xls.addPicture('Network', pic2, 480,735,389,170)
      xls.setCell(32,11,insert_date)
      xls.setCell(47,13,'Daily (5 minutes average)')
      xls.setCell(62,13,'Weekly (30 minutes average)')
      save_close(xls)
def Wed(xls):
      xls.addPicture('Network', pic1, 960,500,389,170)
      xls.addPicture('Network', pic2, 960,735,389,170)
      xls.setCell(32,21,insert_date)
      xls.setCell(47,23,'Daily (5 minutes average)')
      xls.setCell(62,23,'Weekly (30 minutes average)')
      save_close(xls)
def Thu(xls):
      xls.addPicture('Network', pic1, 0,975,389,170)
      xls.addPicture('Network', pic2, 0,1210,389,170)
      xls.setCell(64,1,insert_date)
      xls.setCell(78,3,'Daily (5 minutes average)')
      xls.setCell(94,3,'Weekly (30 minutes average)')
      save_close(xls)
def Fri(xls):
      clearSheet(xls)
      xls = insert_network_Picture(excelPath)
      xls.addPicture('Network', pic1, 0,25,389,170)
      xls.addPicture('Network', pic2, 0,260,389,170)
      xls.setCell(1,1,insert_date)
      xls.setCell(15,3,'Daily (5 minutes average)')
      xls.setCell(30,3,'Weekly (30 minutes average)')
      save_close(xls)
def Sat(xls):
      xls.addPicture('Network', pic1, 480,25,389,170)
      xls.addPicture('Network', pic2, 480,260,389,170)
      xls.setCell(1,11,insert_date)
      xls.setCell(15,13,'Daily (5 minutes average)')
      xls.setCell(30,13,'Weekly (30 minutes average)')
      save_close(xls)
def Sun(xls):
      xls.addPicture('Network', pic1, 960,25,389,170)
      xls.addPicture('Network', pic2, 960,260,389,170)
      xls.setCell(1,21,insert_date)
      xls.setCell(15,23,'Daily (5 minutes average)')
      xls.setCell(30,23,'Weekly (30 minutes average)')
      save_close(xls)
def clearSheet(xls):
      save_close(xls)
      time.sleep(1)
      try:
            wb = openpyxl.reader.excel.load_workbook(excelPath)
            wb.remove_sheet(wb.get_sheet_by_name('Network'))#清空第Network
            wb.create_sheet("Network", 4)
            wb.save(excelPath)
      except Exception(e):
            print(str(e))
        
def save_close(xls):
     xls.save()
     xls.close()   
class insert_network_Picture:  
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
      def addPicture(self, sheet, pictureName, Left, Top, Width, Height):  #插入图片
          sht = self.xlBook.Worksheets(sheet)  
          sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)
      def setCell(self,row,col,value):  #设置单元格的数据
          sht = self.xlBook.Worksheets('Network')
          sht.Cells(row, col).Value = value
                      
if __name__ == "__main__":
      if not os.path.isfile(pic1) or not os.path.isfile(excelPath):
            print('network图片或COSCON-ACZ-daily-stat-result.xlsx不存在')
            sys.exit()
      xls = insert_network_Picture(excelPath)
      result={"Mon":lambda:Mon(xls),
              "Tue":lambda:Tue(xls),
              "Wed":lambda:Wed(xls),
              "Thu":lambda:Thu(xls),
              "Fri":lambda:Fri(xls),
              "Sat":lambda:Sat(xls),
              "Sun":lambda:Sun(xls)
              }
      #result["Fri"]()
      result[ytd]()
      sys.exit()
     
