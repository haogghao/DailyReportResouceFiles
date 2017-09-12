import urllib.request
import ssl
import urllib
import os
import re
import urllib.parse
import time
import datetime

def getTime():
    dt = datetime.datetime.now()
    Ydt = dt + datetime.timedelta(days=-1)
    currentYear = '20'+Ydt.strftime('%y')
    currentMon = Ydt.strftime('%m')
    currentDay = Ydt.strftime('%d')
    DailyFolderName = currentYear+currentMon+currentDay
    return DailyFolderName

def mkdir(path):
    path=path.strip()
    path = path.rstrip('\\')

    isExists = os.path.exists(path)

    if not isExists:
        os.makedirs(path)
        print(path + ' ' + 'is created successful')
        return True
    else:
        print(path + ' ' + 'exists')
        return False

def cbk(a,b,c):
    '''''回调函数
    @a:已经下载的数据块
    @b:数据块大小
    @c:远程文件大小
    '''
    per = 100.0*a*b/c
    if per > 100:
        per = 100
    print('%.2f%%' % per)

def getFiles(FileName,DailyFolderName):
    
    FileName_exchage = urllib.parse.quote(FileName)
    
    getnetwork_url = 'https://hkgwebmail.oocl.com/owa/?ae=Folder&t=IPF.Note&newSch=1&scp=0&id=LgAAAABr%2fGEWSD7oQYjz%2fIZ4wC4zAQDh1Pz7VQdwS6osarko4YdiAAAEaMGvAAAB&slUsng=0&sch=' + FileName_exchage
    response = urllib.request.urlopen(getnetwork_url)
    responseStr = response.read().decode('utf-8')
        
    ttoday = datetime.datetime.now()
    tyear = '20'+ttoday.strftime('%y')
    tmon = ttoday.strftime('%m')
    tday = ttoday.strftime('%d')
    if tmon[0] == '0':
        tmon = tmon[1:]
    else:
        tmon = tmon
    if tday[0] == '0':
        tday = tday[1:]
    else:
        tday = tday
    timestr = tmon+'/'+tday+'/'+tyear
    
    reStr = 'name="chkmsg" (.+?)&nbsp;</td><td nowrap align="right"'
    AfterSelect = re.findall(reStr,responseStr)   
    targetArea = []
    for i in range(len(AfterSelect)):
        if(AfterSelect[i].find(FileName)!=(-1)) and (AfterSelect[i].find(timestr)!=(-1)):
            targetArea.append(AfterSelect[i])
    
    reStr = 'value="(.+?)"'
    EnterDownloadLinkRe_list = []
    EnterDownloadLinkRe_list = re.findall(reStr,str(targetArea))
    url_exchange = urllib.parse.quote_plus(EnterDownloadLinkRe_list[0])   
    
    url_EnterDownload_whole = 'https://hkgwebmail.oocl.com/owa/?ae=Item&t=IPM.Note&id=' + url_exchange
    response = urllib.request.urlopen(url_EnterDownload_whole)
    responseStr = response.read().decode('utf-8')
    
    reStr = '<a id="lnkAtmt" href="(.+?)"'
    DownLoadLinkRe_list = []
    DownLoadLinkRe_list = re.findall(reStr,responseStr)
    
    url_DownLoadFile = url + DownLoadLinkRe_list[0]
    dir = os.path.abspath('.')
    if(FileName == 'COSCON Network Utilization'):
        work_path = os.path.join(dir,'D:\\DailyReportResouceFiles\\'+DailyFolderName+'\\'+FileName+'.zip')
    elif(FileName == 'ACZone Shipment Folder Txn Repor...'):
        work_path = os.path.join(dir,'D:\\DailyReportResouceFiles\\'+DailyFolderName+'\\ACZone Shipment Folder Txn Report.xlsx')
    elif(FileName == 'Coscon User Profile Syn...'):
        work_path = os.path.join(dir,'D:\\DailyReportResouceFiles\\'+DailyFolderName+'\\Report - Coscon User Profile Sync Txn Report.xlsx')
    elif(FileName == 'STDZone COSCON BR SI Daily TXN R...'):
        work_path = os.path.join(dir,'D:\\DailyReportResouceFiles\\'+DailyFolderName+'\\STDZone COSCON BR SI Daily TXN Report.xlsx')
    else:
        work_path = os.path.join(dir,'D:\\DailyReportResouceFiles\\'+DailyFolderName+'\\'+FileName+'.xlsx')
    urllib.request.urlretrieve(url_DownLoadFile,work_path,cbk)
     
ssl._create_default_https_context = ssl._create_unverified_context  #处理证书 //401错误
url = 'https://hkgwebmail.oocl.com/owa/'
passman = urllib.request.HTTPPasswordMgrWithDefaultRealm() #创建域验证对象
passman.add_password(None,url,'cspemon','Password1') #设置域地址，用户名及密码
auth_handler = urllib.request.HTTPBasicAuthHandler(passman) #生成处理与远程主机的身份验证的处理程序
opener = urllib.request.build_opener(auth_handler) #返回一个openerDirector实例
urllib.request.install_opener(opener) #安装一个openerDirector实例作为默认的开启者
response = urllib.request.urlopen(url) #打开URL连接，返回Response对象

DailyFolderName = getTime()
mkpath = 'D:\\DailyReportResouceFiles\\'+DailyFolderName
mkdir(mkpath)

getFiles('COSCON Network Utilization',DailyFolderName)
getFiles('ACZone TXN Monitor',DailyFolderName)
getFiles('ACZone Shipment Folder Txn Repor...',DailyFolderName)
getFiles('Coscon User Profile Syn...',DailyFolderName)
getFiles('STDZone COSCON BR SI Daily TXN R...',DailyFolderName)



