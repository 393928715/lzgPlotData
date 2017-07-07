import win32com.client 
import pymysql
import time
import xlrd
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf8')




class toExcel():
    
    
    def __init__(self):
        self.cursor=self.conlzg().cursor()
        self.loczs=[]
        self.locgp=[]
        self.num=5
        self.loc=[]
        self.filename='E:/test/test.xlsx'
    
    def conlzg(self):
        
        con=pymysql.connect(host='localhost',
        user='root',
        passwd='lzg000',
        db='lzg',
        port=3306,
        charset='utf8'
        )
        return con    
    
    
          
    def buildzs(self):
        start='0'
        left=1    #起始列
        top=1     #起始行
        picleft=1
        pictop=15
        numindex=0
        
        filename='E:\\test\\test.xlsx'
        #wbk =xlsxwriter.Workbook(filename)
        #wbk.close()
        xlApp = win32com.client.Dispatch('Excel.Application') 
        xlBook = xlApp.Workbooks.Open(filename) 
        xlApp.Visible=1
        #xlApp.Sheets.Add()
        sht1=xlBook.Worksheets(1)
        #sht1.name='chart'
        sht2=xlBook.Worksheets(2)
        #xlApp.Sheets("Sheet2").Select
        #sht2.Visible = False
        #sht2.name='data'
        cursor=self.cursor
        cursor.execute('select count(*) from zs group by name')
        s1=cursor.fetchall()
        
        for i in xrange(5):
            num=s1[numindex][0]
            numindex+=1
            count = cursor.execute('select * from zs order by name,time limit'+' '+start+','+str(num))
            self.num=count
            print 'has %s record' % count
            cursor.scroll(0,mode='absolute') 
            results = cursor.fetchall()
            fields = cursor.description    
            for ifs in range(top,len(fields)+top):   
                sht2.Cells(top,ifs+left-top).Value = fields[ifs-top][0]
                
                  
            for row in range(top+1,len(results)+top+1):   
                for col in range(left,len(fields)+left):
                    sht2.Cells(row,col).Value=results[row-1-top][col-left]
                    
            locxy=(row,left)
            self.loc.append(locxy)
                   
            sht1.Shapes.AddChart2(271,4,picleft,pictop,593,200).Select()
            cell1=sht2.cells(top,left+9)
            cell2=sht2.cells(top+99,left+9)
            xlApp.ActiveChart.SetSourceData(Source=sht2.Range(cell1,cell2))
            cell3=sht2.cells(top,left+8)
            cell4=sht2.cells(top,left+8)
            xlApp.ActiveChart.FullSeriesCollection(1).XValues=sht2.Range(cell3,cell4)
            xlApp.ActiveChart.ChartTitle.Text = sht2.cells(top+1,left+2).Value
             
            
            start=int(start)+num
            start=str(start)
            pictop+=200
            top+=100 
        
        #xlBook.Save()
        
    def build(self): 
 
        start=0
        left=14    #起始列
        top=0     #起始行
        pictop=2
        numindex=0

        cursor=self.cursor
        #将字段写入到EXCEL新表的第一行  
        wbk =xlsxwriter.Workbook(self.filename)  
        #newwbk = copy(wbk)
        sheet = wbk.add_worksheet('chart')
        sheet1=wbk.add_worksheet('data')
        #画模块
        data1='指数监控'
        data2='股票监控'
        red = wbk.add_format({'border':4,'align':'center','valign': 'vcenter','bg_color':'C0504D','font_size':16,'font_color':'white'})
        blue = wbk.add_format({'border':4,'align':'center','valign': 'vcenter','bg_color':'8064A2','font_size':16,'font_color':'white'})
        sheet.merge_range(0,0,1,11,data1,red) 
        sheet.merge_range(0,14,1,25,data2,blue)
        
        #每股数量
        cursor.execute('select count(*) from m group by name')
        numlist=cursor.fetchall()
        self.gpnum=numlist[0][0]
        
        for i in xrange(5):
            num=numlist[numindex][0]
            numindex+=1
            
            test = 'select * from m order by name,time limit'+' '+ str(start)+','+ str(num)
            
            cursor.execute('select * from m order by name,time limit'+' '+ str(start)+','+ str(num))
            #print 'has %s record' % count  
            #重置游标位置  
            cursor.scroll(0,mode='absolute')  
            #搜取所有结果  
            results = cursor.fetchall()  
            #获取MYSQL里的数据字段  
            fields = cursor.description  
            #从top行开始写标题            
            #向sheet中插入数据
            for ifs in range( top,len(fields)+ top):  
                sheet1.write( top,ifs+ left- top,fields[ifs- top][0])  
            #写内容    
            for row in range( top+1,len(results)+ top+1):   
                #for col in range(left,len(fields)+left):  
                sheet1.write_row(row, left,results[row-1- top]) 
            
            locxy=[row+1,left+1]
            self.locgp.append(locxy)
            #作图，类型为 line折现图 
            chart1 = wbk.add_chart({'type': 'line'})
            chart1.set_style(4)
            #向图表添加数据 
            chart1.add_series({
            'name':['data', top+1, left+1],
            'categories':['data', top, left+6, top, left+6],
            'values':['data', top+1, left+5, top+200, left+5],
            'line':{'color':'red'},            
            })
            #bold = wbk.add_format({'bold': 1})
            chart1.set_title({'name':'1min Line'}) 
            chart1.set_x_axis({'name':'time',
                                'name_font': {'size': 14, 'bold': True}
                                })
            chart1.set_y_axis({'name':'close',
                               'name_font': {'size': 14, 'bold': True}
                               })
            chart1.set_size({'width':770,'height':300})
            sheet.insert_chart( pictop, left,chart1)
            #bg+=19
            
            start+=num
            top+=300
            pictop+=15
        
        print '股票数据已导入'
        
        
        start=0
        left=0    #起始列
        top=0     #起始行
        pictop=2
        cursor.execute('select count(*) from zs group by name')
        numlist=cursor.fetchall()  
        self.zsnum=numlist[0][0]
        numindex=0
        for i in xrange(5):
            num=numlist[numindex][0]
            numindex+=1
            cursor.execute('select * from zs order by name,time limit'+' '+ str(start)+','+ str(num))
            #print 'has %s record' % count  
            #重置游标位置  
            cursor.scroll(0,mode='absolute')  
            #搜取所有结果  
            results = cursor.fetchall()  
            #获取MYSQL里的数据字段  
            fields = cursor.description  
            #从top行开始写标题            
            #向sheet中插入数据
            for ifs in range( top,len(fields)+ top):  
                sheet1.write( top,ifs+ left- top,fields[ifs- top][0])  
            #写内容    
            for row in range( top+1,len(results)+ top+1):   
                #for col in range(left,len(fields)+left):  
                sheet1.write_row(row, left,results[row-1- top]) 
            
            locxy=(row+1,left+1)
            self.loczs.append(locxy)
            
            #作图，类型为 line折现图 
            chart1 = wbk.add_chart({'type': 'line'})
            chart1.set_style(4)
            #向图表添加数据 
            chart1.add_series({
            'name':['data', top+1, left+2],
            'categories':['data', top, left+6, top, left+6],
            'values':['data', top+1, left+9, top+300, left+9],
            'line':{'color':'red'},            
            'fill':    {'color':'#FF9900'}
            })
            #bold = wbk.add_format({'bold': 1})
            chart1.set_title({'name':'1min line '})
            chart1.set_x_axis({'name':'time',
                                'name_font': {'size': 14, 'bold': True}},                             )
            chart1.set_y_axis({'name':'close',
                               'name_font': {'size': 14, 'bold': True}})
            chart1.set_size({'width':770,'height':300})
            sheet.insert_chart( pictop, left,chart1)
            
            start+=num
            top+=300
            pictop+=15
        
        print '指数数据已导入'
        cursor.close()
        wbk.close()
    
    
    
    
        
    def update(self):
         
        #建立excel工作进程
        xlApp = win32com.client.Dispatch('Excel.Application') 
        xlBook = xlApp.Workbooks.Open(self.filename) 
        xlApp.Visible=1     
        #选定工作sheet
        sht2=xlBook.Worksheets(2)         
        
        while True:       
            S=0
            
            startzs=0
            numindexzs=0
            locindexzs=0   
            
            numindex=0   
            locindex=0
            start=0            
            con=self.conlzg()
            cursor=con.cursor() 
            cursor.execute('select count(*) from zs group by name')
            s1=cursor.fetchall()
        
            for x,y in self.loczs:
                zsnum=s1[numindexzs][0]
                numindexzs+=1
                cursor.execute('select * from zs order by name,time desc limit'+' '+str(startzs)+','+'1')
                if zsnum==self.zsnum:
                    print "指数未更新数据"
                    S=1                   
                    break
                results = cursor.fetchall()   
                for col in range(y,y+10):
                    sht2.Cells(x+1,col).Value=results[0][col-y]
                        
                startzs+=zsnum
    
            if S==0:
                for x,y in self.loczs:
                    self.loczs[locindexzs]=(x+1,y)
                    locindexzs+=1
                self.zsnum+=1
                print '指数数据更新'
                
            
            
              
            S=0
            cursor.execute('select count(*) from m group by name')
            s1=cursor.fetchall()    
            
            for x,y in self.locgp:
                num=s1[numindex][0]
                numindex+=1
                cursor.execute('select * from m order by name,time desc limit'+' '+str(start)+','+'1')
                if num==self.gpnum:           
                    print "股票未更新数据"
                    S=1
                    time.sleep(20)                    
                    break
               
                #cursor.scroll(0,mode='absolute') 
                results = cursor.fetchall()   
                for col in range(y,y+8):
                    sht2.Cells(x+1,col).Value=results[0][col-y]                           
                start+=num
            
            if S==0:
                for x,y in self.locgp:
                    self.locgp[locindex]=(x+1,y)
                    locindex+=1 
                
                self.gpnum+=1
                print'股票数据更新'
                

                cursor.close()
                con.close()        
        

if __name__=='__main__':
    
    
    e=toExcel()       
    e.build()
   # e.update()

