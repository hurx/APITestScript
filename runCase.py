# -*- coding: cp936 -*-
import exclCtl
import httpClass
import os
import json
import ConfigParser

null='null'
true='true'
false='false'

###*************参数构造begin*************####
conf=ConfigParser.ConfigParser()
conf.read('conf.ini')
fileName=conf.get('excel','fileName').decode('cp936')
iSheet=conf.get('excel','iSheet').decode('cp936')
outFile=conf.get('excel','outFile').decode('cp936')+iSheet+("测试结果.html").decode('cp936')
ip=conf.get('server','ip')
port=conf.get('server','port')
###*************参数构造end*************####

exclData=exclCtl.exclCtl(fileName)
nrows=exclData.getRow(iSheet)
if nrows==None:
    print exclData.getTime(),"[ERROR]用例列表为空:1.请检查conf.ini内容是否正确；2.请确认【",iSheet,"】用例是否为空"
    exit()
print "[TOTAL CASE NUMBER]",##打印总的用例个数
print nrows

httpCtl=httpClass.httpRequest(ip,port)

fp=open(outFile,'w')
fp.write('''<HEAD><meta charset="utf-8"></HEAD>''')
fp.write("<BODY>")

for iRow in range(2,nrows+1):
    print ">>>>>>>>>>>>>>>>>>>>>>>>>ROWNUM:",
    print iRow,
    print "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
    try:
        interfaceUrl="/portalzserver"+exclData.readData(iSheet,iRow,4)
        function=exclData.readData(iSheet,iRow,1).decode('utf-8')
        title=exclData.readData(iSheet,iRow,2).decode('utf-8')
        expection=exclData.readData(iSheet,iRow,3).decode('utf-8')
    except:
        print "[ERROR]获取excel表第【",
        print iRow,
        print "】行参数失败,请检查参数格式"
        exit()
    print "[FUNCTION]",
    print function,
    print "[TITLE]",
    print title
    ##print "interfaceUrl="+interfaceUrl
    sendInfo=exclData.readData(iSheet,iRow,5)
    ##print sendInfo
    ##httpCtl.httpGet(interfaceUrl)
    httpRsp=httpCtl.httpPost(sendInfo,interfaceUrl).encode('utf-8')
    result=json.dumps(eval(httpRsp),indent=4).decode('raw_unicode_escape')
    
    #######*********write result into txt begin*******#######
    fp.write('''<span style="color:red;font-weight:bold;">''')
    fp.write(">>>>>>>>>>>>>>>>>>>>>>>>>ROWNUM:"+str(iRow)+"<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<\n")
    fp.write("</span>")
    fp.write("<h4><pre>")
    fp.write("[FUNCTION]:"+function.encode('utf-8'))
    fp.write("    [TITLE]:"+title.encode('utf-8'))
    fp.write("    [EXPECTION]:"+expection.encode('utf-8')+"\n")
    fp.write("</pre></h4>")
    fp.write("<pre>")
    fp.write("[RESULT]:\n")
    fp.write(result.encode('utf-8')+"\n")
    fp.write("</pre>")
    #######*********write result into txt begin*******#######
    ##exclData.writeData(iSheet,iRow,6,httpRsp)
   
fp.write("</BODY>")
httpCtl.httpClose()
exclData.close()
fp.close()
print ''
print "测试结果：",
print outFile
