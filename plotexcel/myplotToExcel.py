# -*- coding: utf-8 -*-
"""
Created on Wed Jun 07 12:09:46 2017

@author: Administrator
"""

from sqlalchemy import create_engine 
import xlsxwriter
import os

import pandas as pd
import numpy as np

import sys
sys.path.append("..\..")
reload(sys)
sys.setdefaultencoding('utf8')

from DataHanle.MktDataHandle import MktIndexHandle

from datetime import datetime


class TmpDealData():
    #初始化程序
    def __init__(self):
        
        #将补齐数据配置到指定文件夹
        self.dbdataDir = '\\BoardIndex\\config\\通达信数据补齐目录\\Tdxdb.txt'       
        
        #一次生成文件个数
        self.fbulidNum = 30
        
         #数据库引擎
        self.engine = create_engine('mysql://root:lzg000@127.0.0.1/stocksystem?charset=utf8')
        
#        
#        #通达信行业目录
#        self.linedir_Day = u'F:\\股票数据\\通达信板块\\行业\\日线'
#        
#        
#        self.linedir_5Min = u'F:\\股票数据\\通达信板块\\行业\\5分钟线'
#        
#        
#        self.linedir_1Min = u'F:\\股票数据\\通达信板块\\行业\\1分钟线'
#        
#        #通达信规模指数目录
#        self.scaleindexdir_Day = u'F:\\股票数据\\通达信板块\\规模指数\\日线'
#        
#        self.scaleindexdir_5Min = u'F:\\股票数据\\通达信板块\\规模指数\\5分钟线'
#        
#        self.scaleindexdir_1Min = u'F:\\股票数据\\通达信板块\\规模指数\\1分钟线'
#        
##        self.scaleindexdir = u'F:\\股票数据\\通达信板块\\规模指数\\test\\'
#        
#        #通达信个股目录
#        self.stockdir_Day = u'F:\股票数据\\个股数据\\日线\\复权\\'
#        
#        self.stockdir_5Min = u'F:\股票数据\\个股数据\\5分钟线\\复权\\'
#        
#        self.stockdir_1Min = u'F:\股票数据\\个股数据\\1分钟线\\复权\\'
    
    #采取读取配置文件方式    
    def getAllTypeDir(self):
        
        #配置文件目录
        
        pwd   =  os.getcwd()
        
        fpwd  = os.path.abspath(os.path.dirname(pwd)+os.path.sep+"..")
        
        dbfname  = fpwd + self.dbdataDir
        
        fheader = ['MktType','DbTable','DataPath']
        #生成目录字典
        fdata=pd.read_table(dbfname.decode('utf-8'),header=1,names=fheader,delimiter=',')
               
        fdatagroup  = fdata.groupby('MktType')
        
        fdict = dict(list(fdatagroup))
        
        #遍历所要补齐的目录
        for fkey in fdict:
                            
            fitemdf = fdict[fkey]
            
            dbtables = fitemdf['DbTable'].tolist()
            
            datapaths = fitemdf['DataPath'].tolist()
            
            #合并成一个dict
            dbtPathdict =dict(zip(dbtables,datapaths)) 
            
            #遍历所有目录
            
            for dbtkey in dbtPathdict:
                
                dtable  = str(dbtkey)
                
                dbInfo = dbtPathdict[dbtkey]
                
                #利用#号来区分时间级别
                dbPos  = dbInfo.find('#')
                
                if  dbPos!=-1:
                    
                    dataPath  = dbInfo[0:dbPos]
                    
                    dataPath  = str(dataPath).decode('unicode_escape')
                    
                    klinetype = dbInfo[dbPos+1:]
                    
                    self.setTdxData(fkey,dataPath,klinetype,self.fbulidNum,dtable,self.engine,2)
                
        m =11
        
    #构建通达信通用数据录入函数
    def setTdxData(self,fkey,fdir,klinetype,fnum,dtable,engine,lastrow):
        
        
        flist = os.listdir(fdir)
        
        sheader_Day =['hq_date','hq_open','hq_high','hq_low','hq_close','hq_vol','hq_amo']
        
        sheader_Min =['hq_date','hq_time','hq_open','hq_high','hq_low','hq_close','hq_vol','hq_amo']
        
        Afdata = pd.DataFrame()
        
        fcount = 0
        
        if flist:
           lastfile=flist[-1]
              
        for fl in  flist:
            
            fname  = fdir + str(fl)   
            
            fcount = fcount +1 
            
            #文件读取第一行信息
            try:
                
                rfile  = open(fname,'r') 
                
                fline = rfile.readline()
                
                flinelist = fline.split()
                
                if len(flinelist)==4 :
                    
                    # 品种代码
                    flcode = flinelist[0]
                    
                    # 品种名称
                    flname = flinelist[1]
                    
                    # 数据类型，通达信支持三种类型数据：1.日线，2. 5分钟， 3. 1分钟
                    flklinetype = flinelist[2]
                     
                    # 数据是否已除权 
                    flrights   = flinelist[3] 
                    
                    flklinetype = flklinetype.strip() 
                    
                    fdataflag = False
                                                            
                    #使用pandas格式读取数据
                                                        
                    if klinetype=="Day":
                                                
                        fdata=pd.read_table(fname,header=1,names=sheader_Day)
                        
                        if len(fdata)>lastrow and lastrow>=1:
                           
                           fdata=fdata.iloc[:-lastrow] 
                                                                              
                           fdatelist =fdata['hq_date'].tolist()
                           
                           fdataflag = True
                           
                        
                    else:
                        
                        fdata=pd.read_table(fname,header=1,names=sheader_Min,dtype={'hq_time':str})
                             
                        if len(fdata)>lastrow and lastrow>=1:
                            
                           fdata=fdata.iloc[:-lastrow] 
                                               
                           fdatedf  =  fdata['hq_date'] + " " + fdata['hq_time']
                                                
                           fdatelist = fdatedf.tolist()
                           
                           fdataflag = True
                    
                    if fdataflag:
                        
                        if '\xef\xbb\xbf'  in flcode:
                            flcode = flcode.replace('\xef\xbb\xbf','')
                            
                        fdcolumnsIndex = fdata.columns
                        
                        # 保留小数点后2位
                        
                        if 'hq_amo' in fdcolumnsIndex:
                            
                            tmp  = fdata['hq_amo'] /100000000
                                                    
                            fdata['hq_amo'] = tmp.round(2)
                            
                        if 'hq_vol' in fdcolumnsIndex:
                            
                            if fkey!='Stock':                       
                                tmp  = fdata['hq_vol'] /10000
                            else:
                                tmp  = fdata['hq_vol'] /1000000
                            
                            fdata['hq_vol'] = tmp.round(2)
                            
                        
                        
                        fdata['hq_code'] = int(flcode)   
                        
                        findex=pd.to_datetime(fdatelist)
                        
                        fdata.set_index(findex,inplace=True)
                        
                        Afdata=Afdata.append(fdata)
                        
                        if fnum<0:
                            fnum  = 1
                        
                        if fcount % int(fnum)==0:
                            
                                Afdata.to_sql(dtable,con=engine,if_exists='append')
                                
                                Afdata=pd.DataFrame()
                        
                        elif fl==lastfile:   
                                            
                            Afdata.to_sql(dtable,con=engine,if_exists='append')
                            
                            Afdata=pd.DataFrame()
                                          
                    m =11
                    
                    print fcount
             
            finally:
                        
                if rfile:            
                    rfile.close()
 
 
class PlotToExcel():
    
    #初始化程序
    def __init__(self):
        
        self.mktindex = MktIndexHandle()     
        
        self.filename = 'E:\\RF-SAP\\PySAP-Master\\PlotData\\export\\excel\\test.xlsx'
        
        pass;
          
    #获取股票数据
    def getPlotStockData(self):
        pass;
    
    #获取指数数据
    def getPlotIndexData(self,indexName,startDate,endDate,KlineType,indexType):
        
        mktindex =  self.mktindex
        
        mktdf = pd.DataFrame()
        
        if indexName!='':
            
            if indexType=='BoardIndex':
                mktdf = mktindex.MktIndexBarHistDataGet(indexName,startDate,endDate,KlineType)
             
            if indexType=='ScaleIndex':
                mktdf = mktindex.MktScaleIndexBarHistDataGet(indexName,startDate,endDate,KlineType)
                
        
        return mktdf
    
        
    #获取涨跌幅
    def getIndexChg(self,idf):
        
        idf_item   = idf.head(1)
        
        idf_tmp    = idf_item['hq_close'].values
        
        idf_close  = idf_tmp[0]
    
        idf['hq_preclose'] = idf['hq_close'].shift(1)
        
        idf['hq_chg']= ((idf['hq_close']/idf['hq_preclose'] -1)*100).round(2)
        
        idf['hq_allchg']= ((idf['hq_close']/idf_close -1)*100).round(2)
        
        idf_ret = idf
        
        return idf_ret
    
    #获取指数相对强弱，相对量能
    def getIndexXdQr(self,bmidf,bkidf,bkdict):
        
        xdidf_ret = pd.DataFrame()
        
        dict_ret  ={}    
        
        #如果板块有数据，计算处理
        
        if len(bkidf)>0 and len(bmidf)>0:
            #数据分组        
            hkidf_group = bkidf.groupby('hq_code')
            
            #生成排名dict        
            hkidf_dict = dict(list(hkidf_group))
            
            #生成基准指数累计涨跌幅度
            xdhidf   = self.getIndexChg(bmidf)
            
            bmi_len   = len(bmidf)
            
            xdhindex  = xdhidf.index
            
            dictcount = 0 
            
            sortdict = {}
    
            #取出排名指数
            for dfdict in  hkidf_dict:
                
                dictcount  = dictcount +1
                
                hidf_item  = hkidf_dict[dfdict]
                
                tmpidf     = xdhidf.copy()
                
                hidf_ret   = self.getIndexChg(hidf_item)
                
                hidf_len   = len(hidf_item)
                
                if hidf_len!=bmi_len:
                    hidf_item = hidf_item.reindex(xdhindex)
               
                tmpdf      = hidf_item.copy()
               
                tmpidf.set_index(['index'], inplace = True)
                
                tmpdf.set_index(['index'], inplace = True)
                
                tmpdf.loc[:,['hq_code']]  = str(dfdict)
                
                tmpdf['hq_bmcode']  = benchmarkIndex 
                
                hq_bmname =''
                
                if bkdict.has_key(str(benchmarkIndex)):
                   hq_bmname = bkdict[str(benchmarkIndex)] 
                
                tmpdf['hq_bmname']  = hq_bmname
                
                
                hq_name =''
                
                if bkdict.has_key(dfdict):
                   hq_name = bkdict[dfdict] 
                
                tmpdf['hq_name']  = hq_name
                
                
                tmpdf.loc[:,['hq_chg']]   = tmpdf['hq_chg']- tmpidf['hq_chg']
                
                tmpdf.loc[:,['hq_allchg']]  = tmpdf['hq_allchg']- tmpidf['hq_allchg']
                
                tmpidf.loc[:,['hq_vol']] = np.where(tmpidf['hq_vol']==0,np.where(tmpdf['hq_vol']==0,-1,-tmpdf['hq_vol']),tmpidf['hq_vol'])
                
                tmpdf['hq_xdvol']  = (tmpdf['hq_vol']/tmpidf['hq_allvol']*100).round(2)
                
                
                tmpcolums = tmpdf.columns
                
                if 'hq_amo' in tmpcolums:                    
                    tmpdf['hq_xdamo']  = (tmpdf['hq_amo']/tmpidf['hq_allamo']*100).round(2)
                
                
                if dictcount<=1:
                    xdhead_array = tmpdf.values
                else:
                    xdhead_array = np.concatenate([xdhead_array,tmpdf.values],axis=0)
                
                
                #加入涨跌幅的排名
                
                hichg = hidf_ret['hq_allchg'].tolist()
                
                if len(hichg)>0:
                   #得到最后一个数据 
                   sortdict[dfdict] =  hichg.pop()
                
                    
                
            xdhead_columnus = tmpdf.columns     
           
            xdhead_idf = pd.DataFrame(xdhead_array,columns=xdhead_columnus)
            
            if 'hq_xdamo' in xdhead_columnus:
                xdidf_ret  = xdhead_idf[['hq_code','hq_name','hq_date','hq_bmcode','hq_bmname','hq_close','hq_preclose','hq_vol','hq_amo','hq_chg','hq_allchg','hq_xdvol','hq_xdamo']]
            
            #对涨跌幅字典进行排序
            dict_ret= sorted(sortdict.items(), key=lambda d:d[1], reverse = True)
        
        return xdidf_ret,dict_ret
        
    
    def bulidChart(self,wbk,data_top,data_left,bkidf_len,bktile,idxstr,shift,data_top2,data_left2,shift2,style):
        
        bk_chart = wbk.add_chart({'type': 'line'})
               
        bk_chart.set_style(4)
       
        #向图表添加数据 
        
               
        if style==1:
            bk_chart.add_series({
             'name':[idxstr, data_top+1, data_left+1],
             'categories':[idxstr, data_top+1, data_left+2, data_top+bkidf_len, data_left+2],
             'values':[idxstr, data_top+1, data_left+shift, data_top+bkidf_len, data_left+shift],
             'line':{'color':'red'},
                    
             })            
            
            bk_chart.add_series({
             'name':[idxstr, data_top2+1, data_left2+1],
             'categories':[idxstr, data_top2+1, data_left2+2, data_top2+bkidf_len, data_left2+2],
             'values':[idxstr, data_top2+1, data_left2+shift2, data_top2+bkidf_len, data_left2+shift2],
             'line':{'color':'blue'},  
             'y2_axis': True,            
             })
        else: 
             
             bk_chart.add_series({
             'name':[idxstr, data_top, data_left+shift-1],
             'categories':[idxstr, data_top+1, data_left+2, data_top+bkidf_len, data_left+2],
             'values':[idxstr, data_top+1, data_left+shift, data_top+bkidf_len, data_left+shift],
             'line':{'color':'FF6347'},
                    
             }) 
                          
             bk_chart.add_series({
             'name':[idxstr, data_top2, data_left2+shift2-1],
             'categories':[idxstr, data_top2+1, data_left2+2, data_top2+bkidf_len, data_left2+2],
             'values':[idxstr, data_top2+1, data_left2+shift2, data_top2+bkidf_len, data_left2+shift2],
             'line':{'color':'708090'},  
             'y2_axis': True,            
             })
        #bold = wbk.add_format({'bold': 1})
     
     
        bk_chart.set_title({'name':bktile,
                           'name_font': {'size': 10, 'bold': True}
                          })
                                   
        bk_chart.set_x_axis({'name_font': {'size': 10, 'bold': True},
                            'label_position': 'low',
                            'interval_unit': 2
                    
                          })
                        
        bk_chart.set_y_axis({'name':u'日期',
                            'name_font': {'size': 10, 'bold': True}
                       })
   
        bk_chart.set_y_axis({'name':'',
                        'name_font': {'size': 10, 'bold': True}
                        })
                       
        bk_chart.set_size({'width':897,'height':300})
           
        
        
        return bk_chart
        
    #构建指数excel构架，,data_left,pic_lef,data_top,pic_top 分别代表数据，图像的 x，y坐标
    
    def bulidExcelPic(self,bkidf_list,wbk,QR_Sheet,Data_Sheet,xdiColumns,data_left,pic_lef,data_top,pic_top):      
           
        #取出排名指数,写入到excel文件中
           
        if len(bkidf_list)>0:
            lastfile = bkidf_list[-1]
           
        for dflist in  bkidf_list:
            
           if len(dflist)==2:
               
               bkidf_code  = dflist[0]
               bkidf_item  = dflist[1]
               
               bkidf_item = bkidf_item.dropna(how='any')
               
               bkhead = bkidf_item.head(1)
               
               bkname = bkhead['hq_name'].values
               
               bkname = bkname[0]
               
               bktile = bkname +'('+bkidf_code+')'
               
               bkidf_item['hq_date'] = bkidf_item['hq_date'].astype('str')
               
               bkidf_len   = len(bkidf_item)
               
               #写入头
               Data_Sheet.write_row(data_top, data_left,xdiColumns)
               
               #写入内容
                   
               for row in range(0,bkidf_len):   
                  #for col in range(left,len(fields)+left):  
                  
                  tmplist  = bkidf_item[row:row+1].values.tolist()
                                              
                  datalist = tmplist[0]
                                                  
                  Data_Sheet.write_row(data_top+row+1, data_left,datalist)
               
               
               idxstr  = u'指数数据'
               
               #指数参数设置
               shift =10
               
               data_top2 = 0
               
               data_left2 = 0
               
               shift2  = 3
               
               style   = 1
               
               bk_chart = self.bulidChart(wbk,data_top,data_left,bkidf_len,bktile,idxstr,shift,data_top2,data_left2,shift2,style)
                    
               #画出双轴对比图
               
               QR_Sheet.insert_chart( pic_top, pic_lef,bk_chart)
                #bg+=19       
             
               pic_lef+=len(xdiColumns)+2
               
               #画成交量与成交金额相对相对量比
               
               shift =11
               
               data_top2 = data_top
               
               data_left2 = data_left
               
               shift2  = 12
               
               style   = 2
               
               bk_chart = self.bulidChart(wbk,data_top,data_left,bkidf_len,bktile,idxstr,shift,data_top2,data_left2,shift2,style)
                    
               #画出双轴对比图
               
               QR_Sheet.insert_chart( pic_top, pic_lef,bk_chart)
                #bg+=19       
                      
               data_top+=bkidf_len +2
               
               pic_top+=15
               
               if lastfile!=dflist: 
                   pic_lef-=len(xdiColumns) +2
               
               
        
        pic_lef+=len(xdiColumns) +2   
        
        return wbk,pic_lef
        
    # 在excel中插入基准指数的数据，lef，top 分别代表 x，y坐标
        
    def bulidIndexDataToExcel(self,bmi_list,Data_Sheet,bmiColumns,left,top):
        
        #写入指数数据头
        Data_Sheet.write_row(top, left,bmiColumns)
                
        for dflist in  bmi_list:
            
           if len(dflist)==2:
                                        
               bkidf_item  = dflist[1]
               
               bkidf_item['hq_date'] = bkidf_item['hq_date'].astype('str')
               
               bkidf_len   = len(bkidf_item)
               
               #写入头
               Data_Sheet.write_row(top, left,bmiColumns)
               
               #写入内容
                   
               for row in range(0,bkidf_len):   
                  #for col in range(left,len(fields)+left):  
                  
                  tmplist  = bkidf_item[row:row+1].values.tolist()
                                              
                  datalist = tmplist[0]
                                                  
                  Data_Sheet.write_row(top+row+1, left,datalist)
        
        return  Data_Sheet  
        
    #构建指数excel构架    
    def bulidIndexExcelFrame(self,bmidf,xdhead_idf,xdtail_idf):
                
        data_left = 0    #数据起始列
        
        pic_left  = 0    #图像起始列 
        
        data_top  = 0    #数据起始行
        
        pic_top   = 2    #图像起始行
        
        wbk =xlsxwriter.Workbook(self.filename)  
        #newwbk = copy(wbk)
        QR_Sheet   = wbk.add_worksheet(u'指数相对强弱')
    
        Data_Sheet = wbk.add_worksheet(u'指数数据')
        
        #画模块
        headStr='指数强弱排名（涨幅）'
        
        headvol= '指数量比（金额）'
        
        tailStr='指数强弱排名（跌幅）'
        
        tailvol= '指数量比（金额）'
        
        
        xdiColumns= list([u'板块代码', u'板块名称', u'日期', u'基准板块代码', u'基准板块名称', u'收盘价', u'前收盘价', u'成交量',u'成交额' ,u'日相对涨跌幅', u'累计相对涨跌幅', u'相对量比', u'相对金额比'])
              
        xdiColumnlens = len(xdiColumns)   
        
        red = wbk.add_format({'border':4,'align':'center','valign': 'vcenter','bg_color':'C0504D','font_size':16,'font_color':'white'})
        
        blue = wbk.add_format({'border':4,'align':'center','valign': 'vcenter','bg_color':'8064A2','font_size':16,'font_color':'white'})
        
        
        #间隔格式
        JG = wbk.add_format({'bg_color':'CCC0DA'})
        
        #处理第一行excel格式
        QR_Sheet.merge_range(0,0,1,xdiColumnlens,headStr,red)         
        QR_Sheet.set_column(xdiColumnlens+1,xdiColumnlens+1,0.3,JG)

        QR_Sheet.merge_range(0,xdiColumnlens+2,1,2*xdiColumnlens+2,tailStr,blue)               
        QR_Sheet.set_column(2*xdiColumnlens+3,2*xdiColumnlens+3,0.3,JG)
                
        QR_Sheet.merge_range(0,2*xdiColumnlens+4,1,3*xdiColumnlens+4,headvol,red)                    
        QR_Sheet.set_column(3*xdiColumnlens+5,3*xdiColumnlens+5,0.3,JG)
        
        QR_Sheet.merge_range(0,3*xdiColumnlens+6,1,4*xdiColumnlens+6,tailvol,blue)
        
        #基准数据写入数据sheet中
        if len(bmidf)>0:
            bmidf_group = bmidf.groupby('hq_code')
    
            bmi_list = list(bmidf_group)
            
            bmiColumns = list([u'基准指数代码', u'基准指数名称', u'日期', u'收盘价', u'前收盘价', u'成交量', u'涨跌幅', u'累涨跌幅',u'总成交量',u'总成交额'])
            
            #未处理多个基准标的比较问题，以及标的指数与板块数据不一致的问题
                  
            Data_Sheet = self.bulidIndexDataToExcel(bmi_list,Data_Sheet,bmiColumns,data_left,data_top)
            
            data_left = data_left +len(bmiColumns) +2
        
        if len(xdhead_idf)>0:
            
            #数据分组        
            xdhead_group = xdhead_idf.groupby('hq_code',sort=False)
            
            bkidf_list= list(xdhead_group)
                        
            (wbk,pic_left)  = self.bulidExcelPic(bkidf_list,wbk,QR_Sheet,Data_Sheet,xdiColumns,data_left,pic_left,data_top,pic_top)
            
            data_left = data_left+xdiColumnlens+2
            
                
        if len(xdtail_idf)>0:            
                       
            #数据分组        
            xdtail_group = xdtail_idf.groupby('hq_code',sort=False)                                           
            #生成dict        
            bkidf_list = list(xdtail_group)
            
            (wbk,pic_left)  = self.bulidExcelPic(bkidf_list,wbk,QR_Sheet,Data_Sheet,xdiColumns,data_left,pic_left,data_top,pic_top)
          
        wbk.close()
                  
    #处理指数数据
    def getExcelIndexData(self,bmidf,bkdict):
        
        retdata  = pd.DataFrame()
        
        if len(bmidf)>0:
            
            tmpdata = bmidf[1:]
            
#            retdata = tmpdata[['hq_code','hq_date','hq_close','hq_preclose','hq_vol','hq_chg','hq_allchg']]
#             
            rethead   = bmidf['hq_code']
#            
            rethcode  = rethead.tolist()
            
            hq_name = ''
            
            if len(rethcode)>0:
                hq_code = rethcode[0]
                
                if(bkdict.has_key(str(hq_code))):                 
                   hq_name = bkdict[str(hq_code)]
                   
            tmpdata['hq_name'] =hq_name        
#            
            retdata = tmpdata[['hq_code','hq_name','hq_date','hq_close','hq_preclose','hq_vol','hq_chg','hq_allchg','hq_allvol','hq_allamo']]
                
        return retdata
        
     # 获取前几，后几排名数据 
    def getSortedIndexdf(self,xd_idf,sortlist,rankingnum):
        
        idf_len = len(sortlist)
        
        headlist = []
        
        taillist = []
        
        bkidf_group = xd_idf.groupby('hq_code')
            
        #生成排名dict        
        bkidf_dict = dict(list(bkidf_group))
        
        xdhead_idf = pd.DataFrame()
        
        xdtail_idf = pd.DataFrame()
        
        #如果个数大于，指定显示数量
        if idf_len>0 and rankingnum<idf_len:
            
             #获取前几后几 
            for sloc in range(rankingnum):
                
                headitem  = sortlist[sloc]
                 
                tailitem  = sortlist[0-sloc-1] 
                
                headlist.append(headitem[0])
                
                taillist.append(tailitem[0])
            
            #取出前几数据
            
            for hlist in headlist:
                
               dictkey = str(hlist) 
               
               if(bkidf_dict.has_key(dictkey)):
                   
                   bkitem = bkidf_dict[dictkey]
                   
                   xdhead_idf=  xdhead_idf.append(bkitem)
                               
            #取出后几数据              
            for hlist in taillist:
                
               dictkey = str(hlist) 
               
               if(bkidf_dict.has_key(dictkey)):
                   
                   bkitem = bkidf_dict[dictkey]
                   
                   xdtail_idf = xdtail_idf.append(bkitem)
                   
        else:
            
           xdhead_idf = xd_idf   
            
        return xdhead_idf,xdtail_idf

        
    #excel中plot相对指数强弱图形
    def PlotIndexPicToExcel(self,benchmarkIndex,allMarketIndex,bkcodestr,startDate,endDate,KlineType,rankingnum,bkdict):
        
        
        indexType = 'BoardIndex'
        
        #获取指数排名数据        
        hidf  =  self.getPlotIndexData(bkcodestr,startDate,endDate,KlineType,indexType)
        
        
        indexType = 'ScaleIndex'
        
        #获取基准指数数据
        bmidf =  self.getPlotIndexData(benchmarkIndex,startDate,endDate,KlineType,indexType)
        
        #获取所有指数成交量，成交额数据        
        allrawdf  =  self.getPlotIndexData(allMarketIndex,startDate,endDate,KlineType,indexType)
        
        #计算所有
        Salldf     =   allrawdf.groupby('hq_date').agg(
                      {'hq_vol':'sum',
                       'hq_amo':'sum'                       
                      })
        
        Salldf['hq_date'] = Salldf.index
        
        Salldf = Salldf.set_index(bmidf.index)
        
        
        bmidf['hq_allvol'] = Salldf['hq_vol']
        
        bmidf['hq_allamo'] = Salldf['hq_amo']
        
        #获取所有排名数据 
        (xd_idf,sortlist) = self.getIndexXdQr(bmidf,hidf,bkdict)
        
        #处理excel中的指数数据
        ebmidf  = self.getExcelIndexData(bmidf,bkdict)
        
        # 获取前几，后几排名数据     
        (xdhead_idf,xdtail_idf) =self.getSortedIndexdf(xd_idf,sortlist,rankingnum)
        
        #画出指数排名图形
        self.bulidIndexExcelFrame(ebmidf,xdhead_idf,xdtail_idf)
        

if '__main__'==__name__:  
    
    pte = PlotToExcel()
    
    tdd = TmpDealData()
    
    dataFlag  = False
     
    # dataFlag  = True
    #调用临时入库程序，完成补齐日线数据   
    if dataFlag:
        tdd.getAllTypeDir()
    
    
    #基准对比指数
    benchmarkIndex = '399317'
        
    benchmarkName  =u'国证A指'
    
    #全市场指数（000002 A股指数，399107深圳A指数）
    allMarketIndex = '000002,399107'
    
    #获取数据起始时间
    
    start_date = datetime.strptime("2017-05-02", "%Y-%m-%d")
    
    end_date = datetime.strptime("2017-06-23", "%Y-%m-%d")
    
    #K线类型    
    KlineType ='D'
        
    #获取所有板块数据
    mktindex  = pte.mktindex
    
    #获取板块与下属关联股票
    (bkLinedf,bkxfdf) = mktindex.MktIndexToStocksClassify('80201')
    
    
    bkcodes = bkLinedf['bz_indexcode'].astype('str')
    
    bkcodelist = bkcodes.tolist()
    
    bkcodestr = ','.join(bkcodes)
    
    tailIndexs=''
    
    rankingnum = 5
    
    bkdict = bkLinedf.set_index('bz_indexcode')['bz_name'].to_dict()
    
    
    
    bkdict[benchmarkIndex]=benchmarkName
    #指数分类对比图
    pte.PlotIndexPicToExcel(benchmarkIndex,allMarketIndex,bkcodestr,start_date,end_date,KlineType,rankingnum,bkdict)
    
    
    m = 1
    
    #所有指数强弱，量比图
    
    
    
    
    
    
    
    