#!/usr/bin/env python   
# -*- coding: utf-8 -*- 
from win32com.client import Dispatch
import win32com.client
import time
class insert_Throughout:  
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
         # "Insert a picture in sheet"  
          sht = self.xlBook.Worksheets(sheet)  
          sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)  
if __name__ == "__main__":
     yesterday= time.strftime("%Y%m%d",time.localtime(time.time()-86400))
     today=time.strftime("%m%d",time.localtime(time.time()))
     BCPath=r'\\'.join(['D:','DailyReportResouceFiles',yesterday,'BC2.png'])
     BLPath=r'\\'.join(['D:','DailyReportResouceFiles',yesterday,'BL2.png'])
     CTPath=r'\\'.join(['D:','DailyReportResouceFiles',yesterday,'CT2.png'])
     excelPath="D:\DailyReportResouceFiles\Report\COSCON-ACZ-daily-stat-result "+today+".xlsx"
     xls = insert_Throughout(excelPath)
     xls.addPicture('Throughout', BCPath, 0,20,1230,270)
     xls.addPicture('Throughout', BLPath, 0,380,1230,270)
     xls.addPicture('Throughout', CTPath, 0,740,1230,270) 
     xls.save()
     xls.close()
