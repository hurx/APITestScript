# -*- coding: cp936 -*-
import exclCtl
import httpClass
import os
import json
import ConfigParser

null='null'
true='true'
false='false'

###*************��������begin*************####
conf=ConfigParser.ConfigParser()
conf.read('conf.ini')
fileName=conf.get('excel','fileName').decode('cp936')
iSheet=conf.get('excel','iSheet').decode('cp936')
outFile=conf.get('excel','outFile').decode('cp936')+iSheet+("���Խ��.html").decode('cp936')
ip=conf.get('server','ip')
port=conf.get('server','port')
###*************��������end*************####

exclData=exclCtl.exclCtl(fileName)
nrows=exclData.getRow(iSheet)
if nrows==None:
    print exclData.getTime(),"[ERROR]�����б�Ϊ��:1.����conf.ini�����Ƿ���ȷ��2.��ȷ�ϡ�",iSheet,"�������Ƿ�Ϊ��"
    exit()
print "[TOTAL CASE NUMBER]",##��ӡ�ܵ���������
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
        print "[ERROR]��ȡexcel��ڡ�",
        print iRow,
        print "���в���ʧ��,���������ʽ"
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
print "���Խ����",
print outFile
