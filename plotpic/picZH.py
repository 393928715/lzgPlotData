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
#识别中文的模块
import sys
reload(sys)
sys.setdefaultencoding('utf8')


class picExcel():
    
    def __init__(self,time,indexChg,stockChg,stock200Chg,indexTldx,sTLDX,sTLDX200,sTLDX30,indexQ,sQ,sQ200,sQ30):
        self.filename=u'E:/工作/报表/综合/综合报表'+time+'.xlsx'
        self.indexChg=indexChg
        self.stockChg=stockChg
        self.stock200Chg=stock200Chg
        
        self.indexTldx=indexTldx
        self.sTLDX=sTLDX
        self.sTLDX200=sTLDX200
        self.sTLDX30=sTLDX30
        
        self.indexQ=indexQ
        self.sQ=sQ
        self.sQ200=sQ200
        self.sQ30=sQ30
        
        self.time=time
    
    
    def picModel(self):
        #新建EXCEL
        wbk =xlsxwriter.Workbook(self.filename)  
        #获得工作sheet
        sheet = wbk.add_worksheet(u'板块统计')
        sheet2=wbk.add_worksheet(u'个股统计')

        #时间模块格式
        blue = wbk.add_format({'font_name':'微软雅黑','border':1,'align':'center','bg_color':'515F85','font_size':11,'font_color':'white'})
        red = wbk.add_format({'border':1,'align':'center','bg_color':'3675A6','font_size':12,'font_color':'white'})
        gray = wbk.add_format({'border':1,'align':'center','bg_color':'F8F8F8','font_size':12,'font_color':'red'})  
        
        #每日标题
        #指数涨幅
        bigQ=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'vcenter','font_size':11,'bg_color':'ABBB8B','font_color':'white'})
        #特立独行
        bigTLDX=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'vcenter','font_size':11,'bg_color':'897A9F','font_color':'white'})        
        #股票涨幅
        bigZF=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'vcenter','font_size':11,'bg_color':'BA8682','font_color':'white'})
        
        #时间间隔
        JG = wbk.add_format({'bg_color':'CCC0DA'})
        #分类间隔
        FLJG = wbk.add_format({'bg_color':'BFBFBF'})
        #正文样式
        ZW=wbk.add_format({'align':'left','font_size':10})
        ZW.set_num_format('0.000')
        TLDX=wbk.add_format({'align':'left','font_size':10})
        PER=wbk.add_format({'align':'left','font_size':10})
        PER.set_num_format('0.00%')
        
        #小标题格式
        ZF=wbk.add_format({'align':'left','font_size':10,'font_color':'FA3858'})
        DF=wbk.add_format({'align':'left','font_size':10,'font_color':'317346'})
        white=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'vcenter','font_size':10})        
        paiming=wbk.add_format({'align':'center','font_size':10})
        #定义写数据的左端和顶端                    
        left=0      
        top=1
        #书写左上角题目
        #sheet.write(0,0,'1min监控',blue)
        
        #写板块统计sheet
        sheet.merge_range(0,left,0,left+24,self.time,red)
                          
        #写每天的标题头    
        sheet.merge_range(top,left,top,left+6,'行业涨跌排序',bigZF)
        sheet.write_row(top+1,left,['行业','涨幅','相对涨幅'],ZF)              
        sheet.write_row(top+1,left+4,['行业','跌幅','相对跌幅'],DF)   
        
        sheet.merge_range(top,left+8,top,left+15,'特立独行',bigTLDX)            
        sheet.write_row(top+1,left+8,['全行业','相关系数'],white)             
        sheet.write_row(top+1,left+11,['行业前5','相关系数'],white)               
        sheet.write_row(top+1,left+14,['行业后5','相关系数'],white) 
        
        sheet.merge_range(top,left+17,top,left+24,'聪明钱',bigQ)
        sheet.write_row(top+1,left+17,['全行业','Q因子'],white)             
        sheet.write_row(top+1,left+20,['行业前5','Q因子'],white)               
        sheet.write_row(top+1,left+23,['行业后5','Q因子'],white)             

        #得到指数涨跌数据
        indexZf=self.indexChg.head(15)
        indexDf=self.indexChg.tail(15)[::-1]
        boardNames=self.indexChg['hq_name']
        
        #得到指数特立独行数据
        indexTldx=self.indexTldx
        indexTldxTop5=indexTldx.head(5)
        indexTldxTail5=indexTldx.tail(5)[::-1]

        #得到指数Q因子数据            
        indexQ=self.indexQ
        indexQTop5=indexQ.head(5)
        indexQTail5=indexQ.tail(5)[::-1]
                  
        #写板块涨幅
        for i in xrange(len(indexZf)): 
            try:
                sheet.write_row(top+i+2,left,[indexZf.iat[i,0],indexZf.iat[i,2],indexZf.iat[i,3]],PER)                                  
            except:
                print indexZf.iat[i,0]    
                
        for i in xrange(len(indexDf)): 
            try:
                sheet.write_row(top+i+2,left+4,[indexDf.iat[i,0],indexDf.iat[i,2],indexDf.iat[i,3]],PER)                                  
            except:
                print indexDf.iat[i,0]          
                                 
        #写板块特立独行
        for i in xrange(len(indexTldx)): 
            try:
                sheet.write_row(top+i+2,left+8,[indexTldx.iat[i,1],indexTldx.iat[i,0]],TLDX)                                  
            except:
                print indexTldx.iat[i,0]        
                
        for i in xrange(len(indexTldxTop5)): 
            try:
                sheet.write_row(top+i+2,left+11,[indexTldxTop5.iat[i,1],indexTldxTop5.iat[i,0]],TLDX)                                  
            except:
                print indexTldxTop5.iat[i,0]                     

        for i in xrange(len(indexTldxTail5)): 
            try:
                sheet.write_row(top+i+2,left+14,[indexTldxTail5.iat[i,1],indexTldxTail5.iat[i,0]],TLDX)                                  
            except:
                print indexTldxTail5.iat[i,0] 
                
                
        #写板块Q因子
        for i in xrange(len(indexQ)): 
            try:
                sheet.write_row(top+i+2,left+17,[indexQ.iat[i,0],indexQ.iat[i,1]],ZW)                                  
            except:
                print indexTldx.iat[i,0]        
                
        for i in xrange(len(indexQTop5)): 
            try:
                sheet.write_row(top+i+2,left+20,[indexQTop5.iat[i,0],indexQTop5.iat[i,1]],ZW)                                  
            except:
                print indexTldxTop5.iat[i,0]                     

        for i in xrange(len(indexQTail5)): 
            try:
                sheet.write_row(top+i+2,left+23,[indexQTail5.iat[i,0],indexQTail5.iat[i,1]],ZW)                                  
            except:
                print indexQTail5.iat[i,0]             
        
                
        #分隔涨幅跌幅
        sheet.set_column(left+3,left+3,0.15)
        #分隔特例独行和板块涨幅
        sheet.set_column(left+7,left+7,1)
        #分隔全行业和行业前5
        sheet.set_column(left+10,left+10,0.15)
        #分隔行业前5和行业后5
        sheet.set_column(left+13,left+13,0.15)
        #分隔特立独行和聪明钱
        sheet.set_column(left+16,left+16,1) 
        #分隔Q全行业和行业前5
        sheet.set_column(left+19,left+19,0.15)        
        #分隔Q前5和后5
        sheet.set_column(left+22,left+22,0.15)        
        #分隔不同时间
       # sheet.set_column(left+25,left+25,0.3,FLJG)             
               

        #写sheet2,个股统计    
        sheet2.merge_range(0,left,0,left+37,self.time,red)
   
        sheet2.merge_range(top,left,top,left+17,'股票涨幅排序',bigZF) 
        sheet2.write_row(top+1,left+1,['市场排名','所有股票','涨幅','相对涨幅','','市场排名','200标的','涨幅','相对涨幅','','市场前20','涨幅','所属','','市场后20','涨幅','所属'],white)            
        sheet2.write(top+1,left,'板块',white)
        sheet2.write(top+1,left+5,'板块',white) 
   
        sheet2.merge_range(top,left+19,top,left+27,'特立独行',bigTLDX) 
        sheet2.write_row(top+1,left+19,['板块','全市场','相关系数','板块','200标的','相关系数','','RF200','RF30'],white)  
        #sheet2.write_row(top+14,left+26,['市场后10','标的后10'],white)
        
        sheet2.merge_range(top,left+29,top,left+37,'聪明钱',bigQ) 
        sheet2.write_row(top+1,left+29,['板块','全市场','Q因子','板块','200标的','Q因子','','RF200','RF30'],white)  
        #sheet2.write_row(top+14,left+36,['市场后10','标的后10'],white)               
            
        #得到股票涨跌数据
        stockChg=self.stockChg
        stockTop=stockChg.head(200)
        stockTail=stockChg.tail(200)[::-1]
        stock200Chg=self.stock200Chg       
        
        RFlen=len(self.sTLDX30)
        tldxTop=self.sTLDX200.head(RFlen)
        #tldxTail=self.sTLDX.tail(10)[::-1]
        tldx200Top=self.sTLDX30
        #tldx200Tail=self.sTLDX200.tail(10)[::-1]
        QTop=self.sQ200.head(RFlen)
        #QTail=self.sQ.tail(10)[::-1]
        Q200Top=self.sQ30
        #Q200Tail=self.sQ200.tail(10)
        
        
        #把特立独行和Q因子前10数据导出到CSV
#        self.sTLDX200.to_csv(u'E:/工作/数据备份/tldx200/'+self.time+'.csv',encoding='gbk')
#        self.sQ200.to_csv(u'E:/工作/数据备份/Q200/'+self.time+'.csv',encoding='gbk')
#        tldx200Top.to_csv(u'E:/工作/数据备份/tldx30/'+self.time+'.csv',encoding='gbk')
#        Q200Top.to_csv(u'E:/工作/数据备份/Q30/'+self.time+'.csv',encoding='gbk')        
        
                  
        #分板块写股票排名  
        top=3

        for board in boardNames:
            data=stockChg[stockChg.board_name==board].head(10)
            dataLen=len(data)
            sheet2.merge_range(top,left,top+dataLen-1,left,board,white)         
            for i in xrange(dataLen):
                try:
                    sheet2.write(top+i,left+1,data.iat[i,5],paiming)
                    sheet2.write_row(top+i,left+2,[data.iat[i,0],data.iat[i,2],data.iat[i,3]],PER)                                  
                except:
                    print data.iat[i,0]                                                                                                  
            top+=dataLen+1
         
        #分板块写200标的排名
        top=3               
        stock200ChgGrouped=stock200Chg.groupby('board_name')
        
        for board,data in stock200ChgGrouped:
           # xdData=stock200XdChg[stock200XdChg.board_name==board]
            dataLen=len(data)
            sheet2.merge_range(top,left+5,top+dataLen,left+5,board,white)
            for i in xrange(dataLen):
                try:
                    sheet2.write(top+i,left+6,data.iat[i,5],paiming)
                    sheet2.write_row(top+i,left+7,[data.iat[i,0],data.iat[i,2],data.iat[i,3]],PER)                                                         
                except:
                    print data.iat[i,1]
                                                           
            top+=(dataLen+1)                 
        
        #写股票前200全市场排名
        top=3
        for i in xrange(len(stockTop)): 
            try:
                sheet2.write_row(top+i,left+11,[stockTop.iat[i,0],stockTop.iat[i,2],stockTop.iat[i,4],'',stockTail.iat[i,0],stockTail.iat[i,2],stockTail.iat[i,4]],PER)                                   
            except:
                print stockTop.iat[i,1]  
        
        #写全市场特立独行
        top=3
        for board in boardNames:
            data=self.sTLDX[self.sTLDX.board_name==board].head(5)
            dataLen=len(data)
            sheet2.merge_range(top,left+19,top+dataLen-1,left+19,board,white)         
            for i in xrange(dataLen):
                try:
                    sheet2.write_row(top+i,left+20,[data.iat[i,2],data.iat[i,0]],TLDX)                                  
                except:
                    print data.iat[i,0]                                                                                                  
            top+=dataLen+1
            
        
        #写200标的特立独行
        top=3               
        stock200TLDXGrouped=self.sTLDX200.groupby('board_name')
        
        for board,data in stock200TLDXGrouped:
           # xdData=stock200XdChg[stock200XdChg.board_name==board]
            dataLen=len(data)
            sheet2.merge_range(top,left+22,top+dataLen,left+22,board,white)
            for i in xrange(dataLen):
                try:
                    sheet2.write_row(top+i,left+23,[data.iat[i,2],data.iat[i,0]],TLDX)                                                         
                except:
                    print data.iat[i,0]
                                                           
            top+=(dataLen+1) 
                  
        #写股票特立独行前10
        top=3
        for i in xrange(len(tldxTop)): 
            try:
                sheet2.write_row(top+i,left+26,[tldxTop.iat[i,2],tldx200Top.iat[i,2]],TLDX)   
                #sheet2.write_row(top+i+13,left+26,[tldxTail.iat[i,2],tldx200Tail.iat[i,2]],TLDX)                                
            except:
                print tldxTop.iat[i,2]  
        
     
                
        #写全市场Q因子
        top=3
        for board in boardNames:
            data=self.sQ[self.sQ.board_name==board]
            data=data.head(5).append(data.tail(3))
            dataLen=len(data)
            sheet2.merge_range(top,left+29,top+dataLen-1,left+29,board,white)         
            for i in xrange(dataLen):
                try:
                    sheet2.write_row(top+i,left+30,[data.iat[i,0],data.iat[i,1]],ZW)                                  
                except:
                    print data.iat[i,0]                                                                                                  
            top+=dataLen+1                
                
        #写200标的Q因子
        top=3               
        stock200QGrouped=self.sQ200.groupby('board_name')
        
        for board,data in stock200QGrouped:
           # xdData=stock200XdChg[stock200XdChg.board_name==board]
            dataLen=len(data)
            sheet2.merge_range(top,left+32,top+dataLen,left+32,board,white)
            for i in xrange(dataLen):
                try:
                    sheet2.write_row(top+i,left+33,[data.iat[i,0],data.iat[i,1]],ZW)                                                         
                except:
                    print data.iat[i,0]
                                                           
            top+=(dataLen+1) 
            

        #写股票Q因子前10
        top=3
        for i in xrange(len(QTop)): 
            try:
                sheet2.write_row(top+i,left+36,[QTop.iat[i,0],Q200Top.iat[i,0]],ZW) 
                #sheet2.write_row(top+i+13,left+36,[QTail.iat[i,0],Q200Tail.iat[i,0]],ZW)                                
            except:
                print QTop.iat[i,0]  
                
        
                
        #分隔200标的与前20
        sheet2.set_column(left+10,left+10,0.2)
        #分隔市场前20后20
        sheet2.set_column(left+14,left+14,0.2)
        #分隔股票涨幅与特立独行
        sheet2.set_column(left+18,left+18,1)
        #分隔全市场特立独行200标的与相关系数
        sheet2.set_column(left+25,left+25,0.3)
        #分隔特立独行和Q因子
        sheet2.set_column(left+28,left+28,1)
        #分隔Q因子200标的与前后10
        sheet2.set_column(left+35,left+35,0.3)

        wbk.close()
        

    def update(self):
        #获得EXCEL进程，以便实时更新    
        xlApp = win32com.client.Dispatch('Excel.Application') 
        xlBook = xlApp.Workbooks.Open(self.filename) 
        xlApp.Visible=1


def main():
    pic=picExcel()
    pic.picModel()
    #pic.update()

if __name__ =='__main__':
    main()
    



