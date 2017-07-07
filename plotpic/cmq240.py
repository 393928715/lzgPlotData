# encoding: utf-8 
#yellow = wbk.add_format({'border':1,'align':'center','bg_color':'C0504D','font_size':12,'font_color':'white'})
#sheet.merge_range(0,0,0,10,data,yellow) 
#sheet.write(1,0,'ab')
#cell=xl_rowcol_to_cell(1, 0,row_abs=True, col_abs=True)
#print cell
"""
Created on Wed Apr 12 10:26:17 2017

@author: Administrator
"""
import xlsxwriter 
#from xlsxwriter.utility import xl_rowcol_to_cell
import pandas as pd
import win32com.client 
import time
#识别中文的模块
import sys
reload(sys)
sys.setdefaultencoding('utf8')


class picExcel():
    
    def __init__(self,timeList,dataList,filename,chgList):
        #时间样本文件目录
        self.timefile=u'E:\\工作\\ExcelWork\\sample.csv'
        #目标生成
        self.filename=u'E:/test/报表/'+filename+'Smart.xlsx'
        self.timeList=timeList
        self.dataList=dataList
        self.chgList=chgList
        self.interval=50       
    
    def picModel(self):
        #新建EXCEL
        wbk =xlsxwriter.Workbook(self.filename)  
        #获得工作sheet
        sheet = wbk.add_worksheet('tldx')

        #设置书写格式
        blue = wbk.add_format({'font_name':'微软雅黑','border':1,'align':'center','bg_color':'515F85','font_size':11,'font_color':'white'})
        red = wbk.add_format({'border':1,'align':'center','bg_color':'3775A9','font_size':12,'font_color':'white'})
        gray = wbk.add_format({'border':1,'align':'center','bg_color':'F8F8F8','font_size':12,'font_color':'red'})  
        #小标题
        bigwhite1=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'vcenter','font_size':11,'bg_color':'B5C398','font_color':'white'})
        bigwhite2=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'vcenter','font_size':11,'bg_color':'C1938F','font_color':'white'})        
        white=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'vcenter','font_size':10})
        #行间隔样式
        JG = wbk.add_format({'bg_color':'CCC0DA'})
        #分类间隔
        FLJG = wbk.add_format({'bg_color':'BFBFBF'})
        #正文样式
        ZW=wbk.add_format({'align':'left','font_size':10})
        #定义写数据的左端和顶端                    
        left=0      
        top=1
        #书写左上角题目
        sheet.write(0,0,'1min监控',blue)
        #书写框架
        num=0
        for t in self.timeList:
            #取得time数据  
            #时间序列颜色交替变换，易于区分           
            if num%2 == 0:
                sheet.merge_range(0,left,0,left+12,t,red)
            else:        
                sheet.merge_range(0,left,0,left+12,t,gray) 
                
            
            #写每天的标题头    
            sheet.merge_range(top,left,top,left+4,'聪明因子排序',bigwhite1)
            sheet.write_row(top+1,left,['全市场','聪明因子'],white)
            
            #sheet.merge_range(top,left+3,top,left+4,'200标的',white)
            sheet.write_row(top+1,left+3,['200标的','聪明因子'],white) 
            
            sheet.merge_range(top,left+6,top,left+12,'涨幅排序',bigwhite2)
            
            sheet.write_row(top+1,left+6,['全市场','涨幅','相对涨幅'],white) 

            sheet.write_row(top+1,left+10,['200标的','涨幅','相对涨幅'],white)             
#            sheet.write_row(top+1,left+12,['全市场','相对涨幅'],white) 
#
#            sheet.write_row(top+1,left+15,['200标的','相对涨幅'],white)            
            
            dataList=self.dataList[0][num].head(10)
            chgList=self.chgList[0][num].head(10)
                    
            try:               
                dataList200=self.dataList[1][num]
            except:
                print t
            try:               
                chgList200=self.chgList[1][num]
            except:
                print t                               
            for i in xrange(len(dataList)): 
                #写全市场的
                try:
                    sheet.write(top+i+2,left,dataList.iat[i,0],ZW)
                    sheet.write(top+i+2,left+1,dataList.iat[i,1],ZW) 
                    sheet.write(top+i+2,left+6,chgList.iat[i,0],ZW)
                    sheet.write(top+i+2,left+7,chgList.iat[i,2],ZW) 
                    sheet.write(top+i+2,left+8,chgList.iat[i,3],ZW)                    
                except:
                    print dataList.iat[i,0]
                
                try:
                    sheet.write(top+i+2,left+3,dataList200.iat[i,0],ZW)                
                    sheet.write(top+i+2,left+4,dataList200.iat[i,1],ZW)
                    sheet.write(top+i+2,left+10,chgList200.iat[i,0],ZW)
                    sheet.write(top+i+2,left+11,chgList200.iat[i,2],ZW)   
                    sheet.write(top+i+2,left+12,chgList200.iat[i,3],ZW)                       
                except:
                    pass
             
#            sheet.set_column(left+2,left+2,0.1,JG)
#            sheet.set_column(left+5,left+5,0.5)
#            sheet.set_column(left+8,left+8,0.1,JG)
#            sheet.set_column(left+11,left+11,0.5)
#            sheet.set_column(left+14,left+14,0.1,JG)  
            sheet.set_column(left+2,left+2,0.15)
            sheet.set_column(left+5,left+5,0.15,FLJG)
            sheet.set_column(left+9,left+9,0.15)
            sheet.set_column(left+13,left+13,0.5,JG)
   
            left+=14       
            num+=1
#        for i in xrange(5):
#            data='测试'            
#            sheet.merge_range(top,0,top+4,0,data,white)
        #top+=54
        #分隔板块
#        sheet.set_row(self.interval+4,1,JG)
#        sheet.set_row(2*self.interval+5,1,JG)
#        sheet.set_row(3*self.interval+6,1,JG)
#        sheet.set_row(4*self.interval+7,1,JG)
#        sheet.set_row(5*self.interval+8,1,JG)
        #sheet.set_column(1,1,0.5)
        wbk.close()
        

    def update(self):
        #获得EXCEL进程，以便实时更新    
        xlApp = win32com.client.Dispatch('Excel.Application') 
        xlBook = xlApp.Workbooks.Open(self.filename) 
        xlApp.Visible=1
        #sht=xlBook.Worksheets(1)
        
#        teststock=u'E:\\股票数据\\测试分钟线\\SH600000.txt'
#        testindex=u'E:\\股票数据\\测试分钟线\\000808.txt'
#        for n in xrange(2):
#            if n ==0:
#                s=pd.read_table(teststock,header=1,usecols=[1,5],names=['time','close'],dtype=str)
#            else:
#                s=pd.read_table(testindex,header=1,usecols=[1,5],names=['time','close'],dtype=str)
#                
#            lens=len(s)
#            for i in xrange(lens):
#                ntime=str(s.iat[i,0])
#                if '9' in ntime:
#                    ntime=ntime.strip('0')
#                    ntime=ntime[0:1]+':'+ntime[1:3]
#                else:    
#                    ntime=ntime[0:2]+':'+ntime[2:4]
#    
#                if ntime in self.timeList:
#                    index=self.timeList.index(ntime)
#                    left=5*index+3
#                    if n ==0:
#                        row=2
#                    else:
#                        row=8
#                    sht.Cells(row,left).Value=s.iat[i,1]
                

        
        
#        while 1:
#            #获取当前时间，刷新数据
#            ntime=(time.ctime())[11:-8]  
#            if ntime in self.timeList:
#                index=self.timeList.index(ntime)
#                left=(index)*4+2
#                sht.Cells(2,left).Value=u'测试'
#                print '写入'
#                time.sleep(30)
#            else:
#               print '未更新'
#               time.sleep(30)


def main():
    pic=picExcel()
    pic.picModel()
    pic.update()

if __name__ =='__main__':
    main()
    



