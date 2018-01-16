#!/usr/local/bin/python2.7
# encoding: utf-8

import wx
import os
import wx.lib.masked as masked
from email.message import Message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.utils import COMMASPACE,formatdate
from email import encoders
import mimetypes
import smtplib 
from email.mime.audio import MIMEAudio
import xlrd
import sys
import imp

imp.reload(sys)
# sys.setdefaultencoding("utf-8")

#初始化主窗口
class Example(wx.Frame):
    global name,hetonggongzi,chanbu,qitayingfa,yingfaheji,queqing,shebao,zhufang,shuiqian,geren,qita,koukuan,yingfabufen,shifajine
    def __init__(self,parent,title):
        super(Example,self).__init__(parent,title=title,size=(500,480),style=wx.DEFAULT_FRAME_STYLE ^(wx.MAXIMIZE_BOX | wx.RESIZE_BORDER))
        self.InitUI()
        self.Centre()
        self.Show()
        
    def InitUI(self):
        menuBar = wx.MenuBar()
        filemenu = wx.Menu()
        helpmenu = wx.Menu()
        menuBar.Append(filemenu,u"&文件")
        fitem = filemenu.Append(1001,u"&导入excel文件","导入excel文件") 
        self.Bind(wx.EVT_MENU, self.OnIn, id=1001)       
        fitem = filemenu.Append(1002,u"&退出系统",u"退出系统")
        self.Bind(wx.EVT_MENU, self.OnExit, id=1002)
        menuBar.Append(helpmenu,u"&帮助")
        helpem = helpmenu.Append(1003,u"&关于",u"关于")
        self.Bind(wx.EVT_MENU, self.OnAbout, id=1003)
        helpem = helpmenu.Append(1004,u"&使用说明",u"使用说明")
        self.Bind(wx.EVT_MENU, self.OnHelp, id=1004)
        self.SetMenuBar(menuBar)
        self.SetTitle(u"data-人力资源-工资发放软件 V7.0.0")
        self.SetBackgroundColour(wx.Colour(100,149,237)) 
        
        panel=wx.Panel(self,-1)
        text=basicLabel = wx.StaticText(panel, -1, u"Excel文件路径",(42,119))
        text2=basicLabel = wx.StaticText(panel, -1, u"工资发放软件V7",(165,60))
    
        font = wx.Font(12, wx.SWISS, wx.NORMAL, wx.BOLD)
        text.SetFont(font)
        font2=wx.Font(17, wx.SWISS, wx.NORMAL, wx.BOLD)
        text2.SetFont(font2)
        self.basicText1 = wx.TextCtrl(panel,-1,size=(200, 25),pos=(170,115))
        self.Center()
        self.Show(True)
        self.button1 = wx.Button(panel, -1, u"发送邮件", pos=(130, 260),size=(80,50))
        self.button1.Bind(wx.EVT_BUTTON, self.OnOpen)
        self.button3 = wx.Button(panel, -1, u"取消发送", pos=(300, 260),size=(80,50))
        self.button3.Bind(wx.EVT_BUTTON, self.OnClear)
        self.button1 = wx.Button(panel, -1, u"...", pos=(390, 115),size=(43,25))
        self.button1.Bind(wx.EVT_BUTTON, self.OnIn)

#导入excel按钮调用的方法        
    def OnIn(self, event):
        dlg1 = wx.FileDialog(self, u"请选择要导入的excel文件...", os.getcwd(),style=wx.OPEN)
        
        if  dlg1.ShowModal() == wx.ID_OK:
            file = dlg1.GetPath()
            self.basicText1.AppendText(file)
#             self.ReadFile()
#             self.SetTitle(self.title + ' -- ' + self.filename)
        dlg1.Destroy()
        
    def OnExit(self, event):        
        self.Close(True)
    def OnAbout(self, event):
        wx.MessageBox(u"Copyright(c)2017-2088 The www\nAll Rights Reserved\nVersion: 7.0.0\nSupport: wuxiaobing@wwww.com", u"关于",wx.ICON_INFORMATION)
    def OnHelp(self, event):
        wx.MessageBox(u"1、该工具适合公司工资邮件群发。\n2、读取文件[excel]格式必须符合规定的模板。\n3、7.0.0版本修改发件人为 wubob@com。", u"帮助",wx.ICON_QUESTION)

#点击发送按钮调用的方法入口处              
    def OnOpen(self,event):
        global biaoti,smtpuser,smtppass,mail,yuefen
        xmlfile=self.basicText1.GetValue()
        if xmlfile=="":
           return wx.MessageBox(u"请选择要发送的EXCEL文件表格。", u"发送信息",wx.ICON_INFORMATION)
        else:
            book = xlrd.open_workbook(xmlfile)
            sheet_name=book.sheet_names()
            table = book.sheets()[0]
            table1 = book.sheets()[1]
            smtpuser=table1.cell(0,0).value
            smtppass=table1.cell(1,0).value
            biaoti=table1.cell(2,0).value
            yuefen=table1.cell(3,0).value
            for i in range(1,table.nrows):
                self.name=table.cell(i,1).value
                mail=table.cell(i,2).value
                self.hetonggongzi=table.cell(i,3).value
                self.chanbu=table.cell(i,4).value
                self.qitayingfa=table.cell(i,5).value
                self.yingfaheji=table.cell(i,6).value
                self.queqing=table.cell(i,7).value
                self.shebao=table.cell(i,8).value
                self.zhufang=table.cell(i,9).value
                self.shuiqian=table.cell(i,10).value
                self.geren=table.cell(i,11).value
                self.qita=table.cell(i,12).value
                self.koukuan=table.cell(i,13).value
                self.yingfabufen=table.cell(i,14).value
                self.shifajine=table.cell(i,15).value
                self.hell(event)
            return  wx.MessageBox(u"恭喜发送完毕，共发送  "+str(i)+u" 位员工的工资条。", u"发送信息",wx.ICON_INFORMATION)
                             
    def OnClear(self,event):
        self.Close(True)

#发送邮件读取表格及正文内容    
    def hell(self,event):
        your_data = {
                'header': ['姓名','合同薪资','餐补','其它应发/扣','应发项目合计','缺勤扣发','社保','住房公积金','税前应发合计','个人所得税','其它扣款','扣款合计','其它应发部分','实发金额'],
                'data': [[str(self.name), str(self.hetonggongzi),str(self.chanbu),str(self.qitayingfa),str(self.yingfaheji),str(self.queqing),str(self.shebao),str(self.zhufang),str(self.shuiqian),str(self.geren),str(self.qita),str(self.koukuan),str(self.yingfabufen),str(self.shifajine)]
                    ]}
                
        zhengfeng =str(' <br />'+self.name+' 您好 :'+' <br />'+
                 '                                              '+' <br />'+
                 '以下是您'+yuefen+'的工资明细:'
                 '                                              '+' <br />'+
                 '特别注意:工资信息属于保密内容，请勿外泄。同时禁止互相讨论工资，一经发现，严肃处理。如有异议，请三日内回复，谢谢您的配合！'+'\n'+
                 '                                              '+' <br />'+
                 '                                      '+'<br />')
        jieshu=str(
            '                                  '+'<br />'+
            '                                  '+'<br />'+
            '                                  '+'<br />'+
            '----------------------------'+'<br />'+
            '武bob'+'<br />'+
            '人力行政部'+'<br />'+
            '座机:80000000-655'+'<br />'+
            '网址:www.data.com'+'<br />'+
            '地址:北京市朝阳区xxxx'+'<br />'+
            '                                  '+'<br />'+
            '<img src="http://fservices.gooagoo.com/android/logo/ddriven.png"/>')
               
        table_template = '<html>'+zhengfeng+'<table border="1" cellspacing="0" cellpadding="1" bordercolor="#000000" width="1800" align="left">{table}</table>'+jieshu+'</html>'
        header_template = '<tr bgcolor="#F79646"><th>{headers}</th></tr>'
        data_template = '<tr align="center"><td>{data}</td></tr>'
        headers = header_template.format(headers='</th><th>'.join(your_data['header']))
        data_str=''
        
        for data in your_data['data']:
            data_str += data_template.format(data='</td><td>'.join(data))    
        html = table_template.format(table=headers + data_str)
        
        
        msg=MIMEText(html,'html','utf-8')
        msg['Subject'] = biaoti
        msg['From'] = smtpuser
        msg['To'] = mail
        msg['Date']= formatdate(localtime=True)
        smtp = smtplib.SMTP()  
        smtp.connect('smtp.data.com')  
        smtp.login(smtpuser, smtppass)
        
     
        try:
            smtp.sendmail(smtpuser, mail, msg.as_string())
            print ( u'发送收件人： '+mail+u'成功 !')
        except Exception:
            print (u'发送收件人： '+mail+u'失败!')
        smtp.quit()

#main主程序函数入口处       
if __name__ == '__main__':
    app = wx.App()
    Example(None,title=u'Layout1')
    app.MainLoop()