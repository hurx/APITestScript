# -*- coding: cp936 -*-
import win32com.client
from win32com.client import Dispatch
import time

class exclCtl:
    def __init__(self,fileName):
        self.xlApp=win32com.client.Dispatch('Excel.Application')
        try:
            self.book=self.xlApp.Workbooks.Open(fileName)
            print self.getTime(),
            print '[LOG]open excel succeed' 
        except:
            print self.getTime(),
            print '[LOG]open excel failed'
            exit()

    def getRow(self,iSheet):
        try:
            sheet=self.book.Worksheets(iSheet)
            nRows=sheet.UsedRange.Rows.Count
            return nRows
        except:
            print self.getTime(),
            print '[LOG]get rows failed'
    
    def getTime(self):
        timeFormat='%Y-%m-%d %X'
        timeNow=time.strftime(timeFormat,time.localtime())
        return timeNow
        
    def close(self):
        self.book.Close()
        self.xlApp.Quit()
        print self.getTime(),
        print '[LOG]excel closed'
        del self

    def readData(self,iSheet,iRow,iCol):
        try:
            sheet=self.book.Worksheets(iSheet)
            dataValue=sheet.Cells(iRow,iCol).Value
            return dataValue.encode('utf-8')
            ##print sValue
        except:
            self.close()
            print self.getTime(),
            print '[LOG]read data failed'
            
    def writeData(self,iSheet,iRow,iCol,sData):
        try:
            sheet=self.book.Worksheets(iSheet)
            sheet.Cells(iRow,iCol).Value=sData.decode('utf-8')
            self.book.Save()
        except:
            self.close()
            print self.getTime(),
            print '[LOG]write data failed'
            exit()


##filename=("E:\\PC测试\\众筹0914\\测试脚本\\test.xlsx").decode('cp936')
##sheet='Sheet1'
##row=1
##col=1
##excel=exclCtl(filename)
##print excel.getTime()
##excel.readData(sheet,row,col)
##excel.writeData(sheet,row,col,'yes')
##excel.close()
