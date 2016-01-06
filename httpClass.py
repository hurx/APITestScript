import httplib,urllib
import time

class httpRequest:
    def __init__(self,IPAdr,IPPort):
        self.hostIP=IPAdr
        self.hostPort=IPPort
        self.headers={"Content-type":"application/json"}
        try:
            self.conn=httplib.HTTPConnection(self.hostIP,self.hostPort,timeout=10)
            print self.getTime(),
            print "[LOG]connect succeed"
        except:
            ##print_eroor_info()
            print self.getTime(),
            print "[LOG]connect fail"
            exit()

    def getTime(self):
        timeFormat='%Y-%m-%d %X'
        timeNow=time.strftime(timeFormat,time.localtime())
        return timeNow

    def httpGet(self,InterfaceUrl):
        self.conn.request("GET",InterfaceUrl)
        response=self.conn.getresponse()
        print response.status,response.reason
        data=response.read()
        ##print data
        return data        
    
    def httpPost(self,params,InterfaceUrl):
        self.conn.request("POST",InterfaceUrl,params,self.headers)
        response=self.conn.getresponse()
        print self.getTime(),
        print "#############Response Begin##################"
        print "[STATUS]",
        print response.status
        print "[REASON]",
        print response.reason
        data=response.read()
        result=data.decode('utf-8')
        ##print "[RSP_INFO]"
        ##print result
        print "####################Response End###################"
        return result

    def httpClose(self):
        self.conn.close()
        print self.getTime(),
        print "[LOG]connect closed"
    

##interface="/portalzserver/ctl/registZcCust"
##ip="10.1.1.87"
##port="8080"
##httpInit=httpRequest(ip,port)

##body='''
##{
##  "sessionContext": {
##    "entityCode": 0,
##    "channel": "",
##    "serviceCode": "",
##    "postingDateText": "",
##    "valueDateText": "",
##    "localDateTimeText": "",
##    "transactionBranch": "",
##    "userId": "",
##    "password": "",
##    "superUserId": "",
##    "superPassword": "",
##    "authorizationReason": "",
##    "externalReferenceNo": "",
##    "userReferenceNumber": "123",
##    "originalReferenceNo": "",
##    "codCustCode": "",
##    "codCustType": "",
##    "codCustCrStatus": "",
##    "codCustBusLimit": "",
##    "mktCode": ""
##  },
##  "codCustType": "1",
##  "codOrgTyp": "",
##  "codRcmCode": "",
##  "codCustUserId": "",
##  "pwdCustLogin": "111111",
##  "codCustPhone": "13811110005",
##  "codVerify": "123456"
##}
##'''
##httpInit.httpPost(body,interface)
##httpInit.httpGet(interface)
##httpInit.httpClose()
