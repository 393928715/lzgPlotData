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
import datetime

#识别中文的模块
import sys
reload(sys)
sys.setdefaultencoding('utf8')


class picExcel():
    
    def __init__(self,datalist,datelist,Qlist):
        #目标生成
        self.filename=u'E:/test/报表/'+datelist[-1]+'tldx.xlsx'
        self.timelist=datelist
        self.datalist=datalist
        self.Qlist=Qlist
        
    
    def picModel(self):                  #,codelist,stocklist,indexlist):
        #新建EXCEL
        wbk =xlsxwriter.Workbook(self.filename)  
        #获得工作sheet
        sheet = wbk.add_worksheet('tldx')
        #设置书写格式
        blue = wbk.add_format({'font_name':'微软雅黑','border':1,'align':'center','bg_color':'515F85','font_size':10,'font_color':'white'})
        #tldx = wbk.add_format({'font_name':'微软雅黑','border':1,'align':'center','font_size':12})
        red = wbk.add_format({'border':1,'align':'center','bg_color':'BE6660','font_size':12,'font_color':'white'})
        gray = wbk.add_format({'border':1,'align':'center','bg_color':'F8F8F8','font_size':12,'font_color':'red'})  
        JG = wbk.add_format({'bg_color':'CDCDCD'}) 
        #sJG= wbk.add_format({'bg_color':'DADADA'})
        #标题格式
        bigwhite1=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'vcenter','font_size':11,'bg_color':'897A9F','font_color':'white'})
        bigwhite2=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'vcenter','font_size':11,'bg_color':'728BAC','font_color':'white'})
        white=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'vcenter','font_size':10})
        index=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'top','font_size':12})
        ZW=wbk.add_format({'align':'center','font_size':10})
        
        #分类间隔
        FLJG = wbk.add_format({'bg_color':'BFBFBF'})     
        
        #定义写数据的左端和顶端                    
        left=0     
        top=1
        #书写左上角题目
        #sheet.write(0,0,'周期监控',blue)
        #书写框架
        num=0
        #按单位时间写每天股票指标排名
        for t in self.timelist:  
            #取得time数据  
            #时间序列颜色交替变换，易于区分
            if num%2 == 0:
                sheet.merge_range(0,left,0,left+20,t,red)
            else:        
                sheet.merge_range(0,left,0,left+20,t,gray)                            
            
            sheet.merge_range(top,left,top,left+9,'特立独行',bigwhite1)
            
            sheet.write(top+1,left,'板块',white)
            sheet.write_row(top+1,left+1,['全市场','相关系数'],white)
            
            #sheet.merge_range(top,left+3,top,left+4,'特立独行',white)
            sheet.write(top+1,left+4,'板块',white)
            sheet.write_row(top+1,left+5,['200标的','相关系数'],white)
            
            sheet.write_row(top+1,left+8,['前20','相关系数'],white)
            
            
            sheet.merge_range(top,left+11,top,left+20,'聪明钱',bigwhite2)
            
            sheet.write(top+1,left+11,'板块',white)
            sheet.write_row(top+1,left+12,['全市场','Q因子'],white) 
            
            #sheet.merge_range(top,left+9,top,left+10,'聪明钱',white)
            sheet.write(top+1,left+15,'板块',white)
            sheet.write_row(top+1,left+16,['200标的','Q因子'],white)    
            
            sheet.write_row(top+1,left+19,['前20','Q因子'],white)              
                           
            #写入数据
            datalist=self.datalist[0][num]
            datatop20=datalist.head(20)
            datalist200=self.datalist[1][num]
            Qlist=self.Qlist[0][num]
            qtop20=(Qlist.head(20)).append(Qlist.tail(20))
            Qlist200=self.Qlist[1][num]
            

            top=3
            stockGrouped=datalist.groupby('board_name')
            for board,data in stockGrouped:
                data=data.head(10)
                dataLen=len(data)
                sheet.merge_range(top,left,top+dataLen,left,board,white)
                for i in xrange(dataLen):
                    try:
                        sheet.write_row(top+i,left+1,[data.iat[i,2],data.iat[i,0]],ZW)                              
                    except:
                        print data.iat[i,1]                                                             
                top+=(dataLen+1)   
                            
          
            top=3
            stock200Grouped=datalist200.groupby('board_name')
            for board,data in stock200Grouped:
                dataLen=len(data)
                sheet.merge_range(top,left+4,top+dataLen,left+4,board,white)
                for i in xrange(dataLen):
                    try:
                        sheet.write_row(top+i,left+5,[data.iat[i,2],data.iat[i,0]],ZW)                               
                    except:
                        print data.iat[i,1]                                                             
                top+=(dataLen+1)   
                
                
            top=3    
            #写前20特立独行
            for i in xrange(len(datatop20)): 
                #写全市场的
                sheet.write_row(top+i,left+8,[datatop20.iat[i,2],datatop20.iat[i,0]],ZW) 
                
#            for j in xrange(len(datalist200)):
#                #写200支标的的
#                sheet.write(top+j+2,left+5,datalist200.iat[j,2],ZW)
#                sheet.write(top+j+2,left+6,datalist200.iat[j,0],ZW)
                

            top=3
            QGrouped=Qlist.groupby('board_name')
            for board,data in QGrouped:
                if len(data)>10:
                    data=(data.head(5)).append(data.tail(5))
                dataLen=len(data)
                sheet.merge_range(top,left+11,top+dataLen,left+11,board,white)
                for i in xrange(dataLen):
                    try:
                        sheet.write_row(top+i,left+12,[data.iat[i,0],data.iat[i,1]],ZW)                               
                    except:
                        print data.iat[i,1]                                                             
                top+=(dataLen+1)                   
                
#            for j in xrange(len(Qlist)):
#                #写2Q因子
#                sheet.write(top+j+2,left+9,Qlist.iat[j,0],ZW)
#                sheet.write(top+j+2,left+10,Qlist.iat[j,1],ZW) 
            top=3
            Q200Grouped=Qlist200.groupby('board_name')
            for board,data in Q200Grouped:
                dataLen=len(data)
                sheet.merge_range(top,left+15,top+dataLen,left+15,board,white)
                for i in xrange(dataLen):
                    try:
                        sheet.write_row(top+i,left+16,[data.iat[i,0],data.iat[i,1]],ZW)                                
                    except:
                        print data.iat[i,1]                                                             
                top+=(dataLen+1)                   
            
            
            top=3
            for i in xrange(len(qtop20)): 
                #写全市场的
                sheet.write_row(top+i,left+19,[qtop20.iat[i,0],qtop20.iat[i,1]],ZW)           
                
            sheet.set_column(left+3,left+3,0.1)
            sheet.set_column(left+7,left+7,0.1)            
            sheet.set_column(left+10,left+10,0.1,JG)
            sheet.set_column(left+14,left+14,0.1)            
            sheet.set_column(left+18,left+18,0.1)
            sheet.set_column(left+21,left+21,0.15,JG)
            
            top=1
            left+=22      
            num+=1
        

        wbk.close()
        

    def update(self):
        #获得EXCEL进程，以便实时更新    
        xlApp = win32com.client.Dispatch('Excel.Application') 
        xlBook = xlApp.Workbooks.Open(self.filename) 
        xlApp.Visible=1
        sht=xlBook.Worksheets(1)
        
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
#                if ntime in self.timelist:
#                    index=self.timelist.index(ntime)
#                    left=5*index+3
#                    if n ==0:
#                        row=2
#                    else:
#                        row=8
#                    sht.Cells(row,left).Value=s.iat[i,1]
                

        
        
#        while 1:
#            #获取当前时间，刷新数据
#            ntime=(time.ctime())[11:-8]  
#            if ntime in self.timelist:
#                index=self.timelist.index(ntime)
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
    



