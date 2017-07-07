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
    
    def __init__(self,filename,timeList,indexChg,stockChg,stock200Chg,stockrelated):
    #stockChg,stock200Chg):
        self.filename=u'E:/test/报表/每日涨幅'+filename+'.xlsx'
        self.timeList=timeList
        self.indexChg=indexChg
        self.stockChg=stockChg
        self.stock200Chg=stock200Chg
        self.stockrelated=stockrelated
        #self.interval=50       
    
    def picModel(self):
        #新建EXCEL
        wbk =xlsxwriter.Workbook(self.filename)  
        #获得工作sheet
        sheet = wbk.add_worksheet('per')

        #时间模块格式
        blue = wbk.add_format({'font_name':'微软雅黑','border':1,'align':'center','bg_color':'515F85','font_size':11,'font_color':'white'})
        red = wbk.add_format({'border':1,'align':'center','bg_color':'3675A6','font_size':12,'font_color':'white'})
        gray = wbk.add_format({'border':1,'align':'center','bg_color':'F8F8F8','font_size':12,'font_color':'red'})  
        #小标题
        bigwhite1=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'vcenter','font_size':11,'bg_color':'ABBB8B','font_color':'white'})
        bigwhite2=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'vcenter','font_size':11,'bg_color':'BA8682','font_color':'white'})        
        white=wbk.add_format({'font_name':'微软雅黑','align':'center','valign':'vcenter','font_size':10})
        #时间间隔
        JG = wbk.add_format({'bg_color':'CCC0DA'})
        #分类间隔
        FLJG = wbk.add_format({'bg_color':'BFBFBF'})
        #正文样式
        ZW=wbk.add_format({'align':'left','font_size':10})
        paiming=wbk.add_format({'align':'center','font_size':10})
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
                sheet.merge_range(0,left,0,left+17,t,red)
            else:        
                sheet.merge_range(0,left,0,left+17,t,gray) 
                          
            #写每天的标题头    
            sheet.merge_range(top,left,top,left+3,'行业涨幅排序',bigwhite1)
            sheet.write_row(top+1,left,['行业','涨幅','相对涨幅'],white)      

            sheet.write(top+1,left+5,'板块',white)
            sheet.write(top+1,left+11,'板块',white)         
            
            sheet.merge_range(top,left+4,top,left+17,'股票涨幅排序',bigwhite2)            
            sheet.write_row(top+1,left+6,['所有股票','涨幅','相对涨幅'],white) 
            #sheet.write_row(top+1,left+10,['全市场','相对涨幅'],white)    
            sheet.write(top+1,left+10,'市场排名',white)
            sheet.write(top+1,left+5,'市场排名',white)
            sheet.write_row(top+1,left+11,['200标的','涨幅','相对涨幅'],white) 
            
            sheet.write_row(top+1,left+15,['前20','涨幅'],white)
            
                       
            #sheet.write_row(top+1,left+17,['200标的','相对涨幅'],white) 
                 
            
            indexChg=self.indexChg[num]
            #indexXdChg=self.indexChg[1][num]
            
            stockChg=self.stockChg[num]
            stockTop20=stockChg.head(300)
            #stockXdChg=self.stockChg[1][num]
            
            stock200Chg=self.stock200Chg[num]
            #stock200XdChg=self.stock200Chg[1][num]
                          
            #写板块排名
            for i in xrange(len(indexChg)): 
                try:
                    sheet.write(top+i+2,left,indexChg.iat[i,0],ZW)
                    sheet.write(top+i+2,left+1,indexChg.iat[i,2],ZW)               
                    sheet.write(top+i+2,left+2,indexChg.iat[i,3],ZW)                                   
                except:
                    print indexChg.iat[i,1]            
                    

            #分板块写股票排名  
            top=3
            
            boardNames=indexChg['hq_name']
            
            for board in boardNames:
                data=stockChg[stockChg.board_name==board].head(10)
                #xdData=stockXdChg[stockXdChg.board_name==board].head(10)
                dataLen=len(data)
                sheet.merge_range(top,left+4,top+dataLen-1,left+4,board,white)
                
                for i in xrange(dataLen):
                    try:
                        sheet.write(top+i,left+5,data.iat[i,5],paiming)
                        sheet.write(top+i,left+6,data.iat[i,0],ZW)
                        sheet.write(top+i,left+7,data.iat[i,2],ZW)
                        #sheet.write(top+i,left+10,xdData.iat[i,0],ZW)                
                        sheet.write(top+i,left+8,data.iat[i,3],ZW)                                   
                    except:
                        print data.iat[i,1]
                                                                                                              
                top+=dataLen+1
                #sheet.set_row(top-1,1,FLJG)                

            boardNames=boardNames.to_frame()
            boardNames['cFlag']=boardNames['hq_name'].isin(stock200Chg['board_name'])
            boardNames=(boardNames[boardNames['cFlag']]).hq_name
 
           
#            top=3
#            for board in boardNames:
#                data200=stock200Chg[stock200Chg.board_name==board]
#                xdData200=stock200XdChg[stock200XdChg.board_name==board]
#                dataLen200=len(data200)
#                sheet.merge_range(top,left+13,top+dataLen200,left+13,board,white)
#                for i in xrange(dataLen):
#                    try:
#                        sheet.write(top+i,left+14,data200.iat[i,0],ZW)
#                        sheet.write(top+i,left+15,data200.iat[i,2],ZW)
#                        sheet.write(top+i,left+17,xdData200.iat[i,0],ZW)                
#                        sheet.write(top+i,left+18,xdData200.iat[i,3],ZW)                                   
#                    except:
#                        print board
#                                                               
#                top+=(dataLen200+1)      
                
                #sheet.set_row(top-1,1,FLJG)
#                
            top=3
#                
            stock200ChgGrouped=stock200Chg.groupby('board_name')
            
            for board,data in stock200ChgGrouped:
               # xdData=stock200XdChg[stock200XdChg.board_name==board]
                dataLen=len(data)
                sheet.merge_range(top,left+9,top+dataLen,left+9,board,white)
                for i in xrange(dataLen):
                    try:
                        sheet.write(top+i,left+10,data.iat[i,5],paiming)
                        sheet.write(top+i,left+11,data.iat[i,0],ZW)
                        sheet.write(top+i,left+12,data.iat[i,2],ZW)
                        #sheet.write(top+i,left+17,xdData.iat[i,0],ZW)                
                        sheet.write(top+i,left+13,data.iat[i,3],ZW)                                   
                    except:
                        print data.iat[i,1]
                                                               
                top+=(dataLen+1)                
#                
        
            top=3
            for i in xrange(len(stockTop20)): 
                try:
                    sheet.write_row(top+i,left+15,[stockTop20.iat[i,0],stockTop20.iat[i,2]],ZW)                                   
                except:
                    print stockTop20.iat[i,1]            
# 
 
    
            sheet.set_column(left+3,left+3,0.15,FLJG)
            sheet.set_column(left+14,left+14,0.15)
            sheet.set_column(left+17,left+17,0.3,JG)
            
            
            top=1
            left+=18  
            num+=1

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
    



