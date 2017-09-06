# -*- coding:utf-8 -*-
import os
import os.path
import zipfile
from zipfile import *
import time

yesterday= time.strftime("%Y%m%d",time.localtime(time.time()-86400))
rootdir = "D:/DailyReportResouceFiles/"+yesterday# 指明压缩文件路径
zipdir = "D:/DailyReportResouceFiles/"+yesterday+"/COSCON Network Utilization"    # 存储解压缩后的文件夹

#Zip文件处理类
class ZFile(object):
    def __init__(self, filename, mode='r', basedir=''):
        self.filename = filename
        self.mode = mode
        if self.mode in ('w', 'a'):
            self.zfile = zipfile.ZipFile(filename, self.mode, compression=zipfile.ZIP_DEFLATED)
        else:
            self.zfile = zipfile.ZipFile(filename, self.mode)
        self.basedir = basedir
        if not self.basedir:
            self.basedir = os.path.dirname(filename)

    def addfile(self, path, arcname=None):
        path = path.replace('//', '/')
        if not arcname:
            if path.startswith(self.basedir):
                arcname = path[len(self.basedir):]
            else:
                arcname = ''
        self.zfile.write(path, arcname)

    def addfiles(self, paths):
        for path in paths:
            if isinstance(path, tuple):
                self.addfile(*path)
            else:
                self.addfile(path)

    def close(self):
        self.zfile.close()

    def extract_to(self, path):
        for p in self.zfile.namelist():
            self.extract(p, path)

    def extract(self, filename, path):
        if not filename.endswith('/'):
            f = os.path.join(path, filename)
            dir = os.path.dirname(f)
            if not os.path.exists(dir):
                os.makedirs(dir)
            open(f, 'wb').write(self.zfile.read(filename))

#解压缩Zip到指定文件夹
def extractZip(zfile, path):
    z = ZFile(zfile)
    z.extract_to(path)
    z.close()

#获得文件名和后缀
def GetFileNameAndExt(filename):
    (filepath,tempfilename) = os.path.split(filename);
    (shotname,extension) = os.path.splitext(tempfilename);
    return shotname,extension

#定义文件处理数量-全局变量
fileCount = 0

#递归获得zip文件集合
def getFiles(filepath):
#遍历filepath下所有文件，包括子目录
  files = os.listdir(filepath)
  for fi in files:
    fi_d = os.path.join(filepath,fi)
    if os.path.isdir(fi_d):
        getFiles(fi_d)
    else:
        global fileCount
        global zipdir
        fileCount = fileCount + 1
        # print fileCount
        fileName = os.path.join(filepath,fi_d)
        filenamenoext = GetFileNameAndExt(fileName)[0]
        fileext = GetFileNameAndExt(fileName)[1]
        # 如果要保存到同一个文件夹，将文件名设为空
        filenamenoext = ""
        zipdirdest = zipdir + "/" + filenamenoext + "/"
        if fileext in ['.zip']:
            if not os.path.isdir(zipdirdest):
                os.mkdir(zipdirdest)
        if fileext == ".zip" :#
            print (str(fileCount) + " -- " + fileName)
           # unzip(fileName,zipdirdest)
            extractZip(fileName,zipdirdest)


#递归遍历“rootdir”目录下的指定后缀的文件列表
getFiles(rootdir)
