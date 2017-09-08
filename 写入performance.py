import csv
import win32com.client
import datetime

class openCSV:
    
    def openFile(self,FileName):
        try:
            self.csv_reader = csv.reader(open(FileName,encoding='utf-8'))
        except Exception(e):
            print(str(e))
        return self.csv_reader

    def getlist(self,csvData):
        self.csvDataList = []
        for row in csvData:
            for col in range(len(row)):
                self.csvDataList.append(row[col])
        return self.csvDataList
    
    def getMaxColNumber(self,csvData):
        self.TableCaption = []
        for row in csvData:
            for col in range(len(row)):
                self.TableCaption.append(row[col])
            break
        return len(self.TableCaption)

class InputExcel:
    def __init__(self,filename=None):
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    '''
    lists:用数组作数据源
    multiplier:乘数，就是原数据每一行的列数
    row_in_sheet:在导入的表里开始位置的行坐标
    col_in_sheet:在导入的表里开始位置的列坐标
    sheet:创的表格
    '''
    def setCell(self,lists,multiplier,row_in_sheet,col_in_sheet,sheet):
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
        return rows

    def setDateCell(self,theLastRowPos,sheet):
        self.sht = self.xlBook.Worksheets(sheet)
        self.dt = datetime.datetime.now()
        self.Ydt = self.dt + datetime.timedelta(days = -1)
        self.Ytyear = '20'+self.Ydt.strftime("%y")
        self.Ytmon = self.Ydt.strftime("%m")
        self.Ytday = self.Ydt.strftime("%d")
        self.timeStr = self.Ytyear+'-'+self.Ytmon+'-'+self.Ytday
        self.d = self.sht.Cells(theLastRowPos+1,1)
        self.d.Value = self.timeStr+' 00:00:00  to '+self.timeStr+' 23:59:59 HKT'

    def getTheLastRowPos(self,sheet):
        self.sht = self.xlBook.Worksheets(sheet)
        self.number = 1
        while 1:
            self.cell = self.sht.Cells(self.number,1).value
            self.cellNext = self.sht.Cells(self.number+2,1).value
            if self.cell or self.cellNext:
                self.number = self.number + 1
            else:
                break
        return self.number

    def save(self,newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()
    def close(self):
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp
    

if __name__ == '__main__':

    dt = datetime.datetime.now()
    dd = dt + datetime.timedelta(days = -1)
    ydt = dt + datetime.timedelta(days = -2)
    sourceFileKeyword = '20'+ydt.strftime("%y")+ydt.strftime("%m")+ydt.strftime("%d")
    targetFilekeyword = dd.strftime("%m")+dd.strftime("%d")
    
    opener = openCSV()
    
    ##获取COSCONACZ的数据
    csvData_cosconacz = opener.openFile("D:\\DailyReportResouceFiles\\"+sourceFileKeyword+"\\CS2-ACZ-COSCONACZ-PROD.csv")
    csvDataList_cosconacz = opener.getlist(csvData_cosconacz)
    csvData_cosconacz = opener.openFile("D:\\DailyReportResouceFiles\\"+sourceFileKeyword+"\\CS2-ACZ-COSCONACZ-PROD.csv")
    multiplier_cosconacz = opener.getMaxColNumber(csvData_cosconacz)
    ##获取COSCON的数据
    csvData_coscon = opener.openFile("D:\\DailyReportResouceFiles\\"+sourceFileKeyword+"\\CS2-ACZ-COSCON-PROD.csv")
    csvDataList_coscon = opener.getlist(csvData_coscon)
    csvData_coscon = opener.openFile("D:\\DailyReportResouceFiles\\"+sourceFileKeyword+"\\CS2-ACZ-COSCON-PROD.csv")
    multiplier_coscon = opener.getMaxColNumber(csvData_coscon)

    xls = InputExcel(r"D:\\DailyReportResouceFiles\\Report\\COSCON-ACZ-daily-stat-result "+targetFilekeyword+".xlsx")
    theLastRowPos = xls.getTheLastRowPos('ServerPerformance')
    xls.setDateCell(theLastRowPos,'ServerPerformance')
    
    if int(multiplier_cosconacz) ==  int(multiplier_coscon) or int(multiplier_cosconacz) > int(multiplier_coscon):
        acz_row = xls.setCell(csvDataList_cosconacz,multiplier_cosconacz,theLastRowPos+2,1,'ServerPerformance')
        coscon_row = xls.setCell(csvDataList_coscon[multiplier_coscon:],multiplier_coscon,theLastRowPos+2+acz_row,1,'ServerPerformance')

    else:
        coscon_row = xls.setCell(csvDataList_coscon,multiplier_coscon,theLastRowPos+2,1,'ServerPerformance')
        acz_row = xls.setCell(csvDataList_cosconacz[multiplier_cosconacz:],multiplier_cosconacz,theLastRowPos+2+coscon_row,1,'ServerPerformance')
    xls.save()
    xls.close()
    
