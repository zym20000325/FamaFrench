import os
import pandas as pd
import numpy as np
import xlwt
from statsmodels.formula.api import ols





##################################### 数据导入处理和因子指标计算部分 #####################################


"""
Created on Wed Nov 25 02:37:40 2020

@author: zym
"""



Cmretwdos = []  # 考虑现金红利再投资的综合月市场回报率(流通市值加权平均法)
Nrrmtdt = []    # 月度化无风险利率(%)



# 读取财务数据
# 从 2014-03-31 到 2017-03-31

print("\nStart to read balance sheet and income statement \n")

Data_finance = []        # 存储读取的所有财务数据  

y = 2013

for i in range(13):

    data_tmp = []
    data_tmp.clear()

    if(i%4==0):
        y = y+1       # 2014 2015 2016 2017
    m = (i%4+1)*3     # 3 6 9 12

    data_tmp.append(y)
    data_tmp.append(m)

    if(m<=9):
        period = str(y)+"-0"+ str(m)
    else:
        period = str(y)+"-"+ str(m)
    fileName = os.getcwd() + "\\财务数据\\" + period+".xlsx"
    print(period)
    df = pd.DataFrame(pd.read_excel(fileName))
    data_tmp.append(df)

    #print(data_tmp[0])
    #print(data_tmp[1])
    #print(len(data_tmp[2]))

    Data_finance.append(data_tmp)

# Data_finance 结构为 13×3    Data_finance[i][0] 为年份   Data_finance[i][1] 为月份   Data_finance[i][2] 为对应时间的财务数据 
# 这样的格式记录年份和月份 方便后续查找
      
print("\nFinished\n")



# 读取市场回报数据
# 2014.6 ~ 2017.6

print("\nStart to read stockmnth \n ")

Data_stock = []   # 存储读取的所有市场回报数据

y = 2014

for i in range(37):

    #stocks_tmp = []
    data_tmp = []
    Cmretwdos_tmp = []
    Nrrmtdt_tmp = []
    #stocks_tmp.clear()
    data_tmp.clear()
    Cmretwdos_tmp.clear()
    Nrrmtdt_tmp.clear()

    if((i+6)%12==1):
        y = y+1
    m = (i+6)%12
    if(m==0):
        m=12

    data_tmp.append(y)
    data_tmp.append(m)

    if(m<=9):
        period = str(y)+"-0"+ str(m)
    else:
        period = str(y)+"-"+ str(m)

    
    fileName = os.getcwd() + "\\市场回报数据\\"+period+".xlsx"
    print(period)

    Cmretwdos_tmp.append(period)
    Nrrmtdt_tmp.append(period)

    df = pd.DataFrame(pd.read_excel(fileName))
    
    data_tmp.append(df)
    Cmretwdos_tmp.append(df.Cmretwdos[0])
    Nrrmtdt_tmp.append(df.Nrrmtdt[0])

    #print(data_tmp[0])
    #print(data_tmp[1])
    #print(len(data_tmp[2]))

    Data_stock.append(data_tmp)
    Cmretwdos.append(Cmretwdos_tmp)
    Nrrmtdt.append(Nrrmtdt_tmp)

# Data_stock 结构为 37×3    Data_stock[i][0] 为年份   Data_stock[i][1] 为月份   Data_stock[i][2] 为对应时间的财务数据 

print("\nFinished  \n ")





# define_stock 函数的作用是取可用股票的交集 因为在每一个周期中需要用到的股票信息包括本月 上个月 上个季度及上上个季度 需要取交集 确保股票信息不缺失
# 具体的方法是先根据当前周期确定年份和月份 算出上个月 上个季度及上上个季度的具体年份和月份 
# 在 Data_finance 和 Data_stock 中查找对应时间的所有股票的 Stkcd 然后取交集

def define_stock(t):
    
    this_year = t[0]   
    this_month = t[1]   
 
    last_year = t[0]-1   
    last_month = t[1]-1   
    if(last_month==0):
        last_month=12
    
    s = (this_month-1)/3   
    s = int(s)
    last_season = s*3     

    if(last_season==0):
        last_season=12
    
    last_last_season = last_season-3    
    if(last_last_season==0):
        last_last_season=12

    #print("this_year "+str(this_year))
    #print("this_month "+str(this_month))

    #print("last_year "+str(last_year))
    #print("last_month "+str(last_month))

    #print("last_season "+str(last_season))
    #print("last_last_season "+str(last_last_season))
    
    stocks = []
    stocks_this_month = []  
    stocks_last_month = []
    stocks_last_season_s = []

    for i in range(len(Data_stock)):

        # 这个月 stock 
        
        if(Data_stock[i][0]==this_year and Data_stock[i][1]==this_month):

            l = len(Data_stock[i][2])
            for j in range(l):

                stocks.append(Data_stock[i][2].Stkcd[j])
                stocks_this_month.append(Data_stock[i][2].Stkcd[j])
            
        # 上个月 stock

        if(last_month==12):      # 如果上个月是12月 年份应为去年
            y=last_year
        else:
            y=this_year
        
        if(Data_stock[i][0]==y and Data_stock[i][1]==last_month):

            l = len(Data_stock[i][2])
            for j in range(l):

                stocks_last_month.append(Data_stock[i][2].Stkcd[j])
        
        # 上个季度 stock in Data_stock

        if(last_season==12):
            y=last_year
        else:
            y=this_year
        
        if(Data_stock[i][0]==y and Data_stock[i][1]==last_season):

            l = len(Data_stock[i][2])
            for j in range(l):

                stocks_last_season_s.append(Data_stock[i][2].Stkcd[j])

    stocks_last_season_f = []
    stocks_last_last_season = []

    for i in range(len(Data_finance)):

        # 上个季度 stock in Data_finance

        if(last_season==12):
            y=last_year
        else:
            y=this_year
        
        #if(i==0):
            #print(y)
            #print(last_season)

        if(Data_finance[i][0]==y and Data_finance[i][1]==last_season):

            l = len(Data_finance[i][2])
            for j in range(l):

                stocks_last_season_f.append(Data_finance[i][2].Stkcd[j])

        # 上上个季度 stock in Data_finance

        if(last_last_season==12 or last_last_season==9):
            y=last_year
        else:
            y=this_year
        
        #if(i==0):
            #print(y)
            #print(last_last_season)

        if(Data_finance[i][0]==y and Data_finance[i][1]==last_last_season):

            l = len(Data_finance[i][2])
            for j in range(l):

                stocks_last_last_season.append(Data_finance[i][2].Stkcd[j])


    # 取交集

    #print(len(stocks_this_month))
    #print(len(stocks_last_month))
    #print(len(stocks_last_season_s))
    #print(len(stocks_last_season_f))
    #print(len(stocks_last_last_season))

    stocks = list(set(stocks_this_month).intersection(set(stocks))) 
    stocks = list(set(stocks_last_month).intersection(set(stocks))) 
    stocks = list(set(stocks_last_season_s).intersection(set(stocks))) 
    stocks = list(set(stocks_last_season_f).intersection(set(stocks))) 
    stocks = list(set(stocks_last_last_season).intersection(set(stocks))) 

    #print(len(stocks))

    return stocks, this_year, this_month, last_year, last_month,last_season,last_last_season


# Mretwd_p 函数的作用是计算一个投资组合S的组合收益率（Mretwd） 按照 Size 加权平均

def Mretwd_p(S):

    n = len(S)

    Size_S = 0
    for i in range(n):
        Size_S = Size_S + S[i]['Size']

    result = 0

    for i in range(n):
        result = result + S[i]['Mretwd']*S[i]['Size']/Size_S
        #result = result + S[i]['Mretwd']

    #result = result/n
    return result


# Divide_Stock 函数的作用是排序和分组
# 首先将股票按照某一个指标从小到大排序  t='Size' 'BM' 'OP' 或 'Inv'
# 'Size' 分两组   'BM' 'OP' 'Inv' 分三组

def Divide_Stock(S,t):

    # 排序
    Stock1 = sorted(S, key=lambda k: k[t])     

    if(t=='Size'):     

        t1 = int(len(Stock1)/2)
        t2 = len(Stock1)-t1
        Stock_S = []
        for i in range(t1):
            Stock_S.append(Stock1[i])
        Stock_B = []
        for i in range(t2):
            Stock_B.append(Stock1[i+t1])
        return Stock_S, Stock_B

    else:

        a1 = int(len(Stock1)*0.3)
        a2 = int(len(Stock1)*0.7)
        a3 = len(Stock1)-a2

        Stock_11 = []
        for i in range(a1):
            Stock_11.append(Stock1[i])     # 0 a1-1

        Stock_12 = []
        for i in range(a2-a1):
            Stock_12.append(Stock1[i+a1])   # a1 a2-1

        Stock_13 = []
        for i in range(a3):
            Stock_13.append(Stock1[i+a2])   # a2 a2+a3-1=len(Stock11)-1

        return Stock_11, Stock_12, Stock_13


# 将一个投资组合S及相关指标写入EXCEL

def write_Excel(S,name,period):

    filepath = (os.getcwd()+"\\分组\\"+period) 
    isExists=os.path.exists(filepath)
    if not isExists:
        os.makedirs(filepath)
    excelpath = (os.getcwd()+"\\分组\\"+period+"\\"+name+".xls") 
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet('Sheet1',cell_overwrite_ok=True)
    headlist=['Stkcd','Size','BM','OP','Inv','Mretwd']
    row=0
    col=0
    for head in headlist:
        sheet.write(row,col,head)
        col=col+1
    n = len(S)
    for i in range(n):
        sheet.write(i+1,0,int(S[i]['Stkcd']))
        sheet.write(i+1,1,float(S[i]['Size']))
        sheet.write(i+1,2,float(S[i]['BM']))
        sheet.write(i+1,3,float(S[i]['OP']))
        sheet.write(i+1,4,float(S[i]['Inv']))
        sheet.write(i+1,5,float(S[i]['Mretwd']))

    workbook.save(excelpath) 

# 将计算所得因子写入EXCEL

def write_Excel_f(results):

    filepath = (os.getcwd()+"\\结果") 
    isExists=os.path.exists(filepath)
    if not isExists:
        os.makedirs(filepath)
    excelpath = (os.getcwd()+"\\结果\\"+"results.xls") 
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet('Sheet1',cell_overwrite_ok=True)
    headlist=['period','SMB','HML','RMW','CMA']
    row=0
    col=0
    for head in headlist:
        sheet.write(row,col,head)
        col=col+1
    
    n = len(results)
    for i in range(n):
        sheet.write(i+1,0,str(results[i][0]))
        sheet.write(i+1,1,float(results[i][1]))
        sheet.write(i+1,2,float(results[i][2]))
        sheet.write(i+1,3,float(results[i][3]))
        sheet.write(i+1,4,float(results[i][4]))
    
    workbook.save(excelpath) 



period_all = []

for year in [2014,2015,2016,2017]:
    for month in [1,2,3,4,5,6,7,8,9,10,11,12]:
        period_tmp = []
        period_tmp.clear()
        period_tmp.append(year)
        period_tmp.append(month)
        period_all.append(period_tmp)

period_all = period_all[5:42]
#print(period_all)


# 记录每次循环所得 SMB HML RMW CMA
# 第i次循环   results[i][0]为period   results[i][1]为SMB   results[i][2]为HML   results[i][3]为RMW   results[i][4]为CMA

results = []    


# 开始循环  2014.7 ~ 2017.6

print("\nStart to calculate SMB HML RMW CMA\n")

for p in range(len(period_all)):

    if(p==0):
        continue    # 跳过2014.6 

    t= period_all[p]
    #t = [2017,6]

    period = str(t[0])+"-"

    if(t[1]<=9):
        period = period+"0"+str(t[1])
    else:
        period = period+str(t[1])

    stocks,this_year, this_month, last_year, last_month,last_season,last_last_season = define_stock(t)

    # print(period+ "  num of stocks: "+str(len(stocks))+"\n")
    print(period)

    
    # 根据交集的stocks读出可用的财务数据和市场回报数据


    # t 月的市场回报
    mretwd_this_month = []

    # t 月的流通市值
    msmvosd_this_month = []

    # t-1 月的流通市值
    msmvosd_last_month = []

    # 上个季度的流通市值
    msmvosd_last_season = []

    m2=last_month
    if(m2==12):
        y2=last_year
    else:
        y2=this_year

    # print(y2)
    # print(m2)

    m3=last_season
    if(m3==12):
        y3=last_year
    else:
        y3=this_year

    # print(y3)
    # print(m3)

    for i in range(len(Data_stock)):

        # t 月的市场回报
        # t 月的流通市值

        if(Data_stock[i][0]==this_year and Data_stock[i][1]==this_month):

            df = Data_stock[i][2]
            l = len(df)

            for j in range(l):

                tmp = []
                tmp1 = []
                tmp.clear()
                tmp1.clear()

                if(df.Stkcd[j] in stocks):

                    tmp.append(df.Stkcd[j])
                    tmp.append(df.Mretwd[j])
                    mretwd_this_month.append(tmp)

                    tmp1.append(df.Stkcd[j])
                    tmp1.append(df.Msmvosd[j])
                    msmvosd_this_month.append(tmp1)

        # t-1 月的流通市值
                    
        if(Data_stock[i][0]==y2 and Data_stock[i][1]==m2):

            df = Data_stock[i][2]
            l = len(df)

            for j in range(l):

                tmp2 = []
                tmp2.clear()

                if(df.Stkcd[j] in stocks):

                    tmp2.append(df.Stkcd[j])
                    tmp2.append(df.Msmvosd[j])
                    msmvosd_last_month.append(tmp2)

        # 上个季度的流通市值

        if(Data_stock[i][0]==y3 and Data_stock[i][1]==m3):

            df = Data_stock[i][2]
            l = len(df)

            for j in range(l):

                tmp3 = []
                tmp3.clear()

                if(df.Stkcd[j] in stocks):

                    tmp3.append(df.Stkcd[j])
                    tmp3.append(df.Msmvosd[j])
                    msmvosd_last_season.append(tmp3)


    # 上个季度的账面价值 股东权益合计
    total_equity_last_season = []

    # 上个季度的营业利润
    operating_profit_last_season = []

    # 上个季度的总资产
    total_assets_last_season = []

    # 上上个季度的总资产
    total_assets_last_last_season = []

    m4 = last_last_season
    if(m4==12 or m4==9):
        y4=last_year
    else:
        y4=this_year

    for i in range(len(Data_finance)):

        # 上个季度的账面价值/股东权益合计
        # 上个季度的营业利润
        # 上个季度的总资产

        if(Data_finance[i][0]==y3 and Data_finance[i][1]==m3):

            df = Data_finance[i][2]
            l = len(df)

            for j in range(l):

                tmp = []
                tmp1 = []
                tmp2 = []
                tmp.clear()
                tmp1.clear()
                tmp2.clear()

                if(df.Stkcd[j] in stocks): 

                    tmp.append(df.Stkcd[j])
                    tmp.append(df.total_equity[j]) 
                    total_equity_last_season.append(tmp)

                    tmp1.append(df.Stkcd[j])
                    tmp1.append(df.operating_profit[j]) 
                    operating_profit_last_season.append(tmp1)

                    tmp2.append(df.Stkcd[j])
                    tmp2.append(df.total_assets[j]) 
                    total_assets_last_season.append(tmp2)


        # 上上个季度的总资产

        if(Data_finance[i][0]==y4 and Data_finance[i][1]==m4):

            df = Data_finance[i][2]
            l = len(df)

            for j in range(l):

                tmp3 = []
                tmp3.clear()

                if(df.Stkcd[j] in stocks): 

                    tmp3.append(df.Stkcd[j])
                    tmp3.append(df.total_assets[j]) 
                    total_assets_last_last_season.append(tmp3)


    # print(len(stocks))
    # print(stocks[300])

    # print(len(mretwd_this_month))
    # print(mretwd_this_month[300][0])
    # print(mretwd_this_month[300][1])

    # print(len(msmvosd_this_month))
    # print(msmvosd_this_month[300][0])
    # print(msmvosd_this_month[300][1])

    # print(len(msmvosd_last_month))
    # print(msmvosd_last_month[300][0])
    # print(msmvosd_last_month[300][1])

    # print(len(msmvosd_last_season))
    # print(msmvosd_last_season[300][0])
    # print(msmvosd_last_season[300][1])

    # print(len(total_equity_last_season))
    # print(total_equity_last_season[300][0])
    # print(total_equity_last_season[300][1])

    # print(len(operating_profit_last_season))
    # print(operating_profit_last_season[300][0])
    # print(operating_profit_last_season[300][1])

    # print(len(total_assets_last_season))
    # print(total_assets_last_season[300][0])
    # print(total_assets_last_season[300][1])

    # print(len(total_assets_last_last_season))
    # print(total_assets_last_last_season[300][0])
    # print(total_assets_last_last_season[300][1])

    

    Stock_ALL = []     

    # Stock_ALL 记录在当前周期股票及相关指标 
    # 投资组合的格式是一个包含若干dict的list 每一个dict代表一个股票
    # dict的格式形如 {'Stkcd': 819, 'Size': 4080142.22, 'BM': 119.2730191790299, 'OP': 0.0053602077197069786, 'Inv': 0.00967669309353584, 'Mretwd': 0.129908}


    for i in range(len(stocks)):

        one_stock = {}
        one_stock.clear()

        # Stkcd  股票序号
        one_stock['Stkcd'] = stocks[i]

        # Size   股票第t-1月的流通市值 Msmvosd
        one_stock['Size'] = msmvosd_last_month[i][1]

        # BM     上个季度的账面价值/流通市值 total_equity / Msmvosd
        one_stock['BM'] = total_equity_last_season[i][1]/msmvosd_last_season[i][1]

        # OP     上个季度的营业利润/股东权益合计 operating profit / total_equity
        one_stock['OP'] = operating_profit_last_season[i][1]/total_equity_last_season[i][1]

        # Inv    [上个季度 total_assets - 上上个季度 total_assets]  /  上上个季度 total_assets
        one_stock['Inv'] = (total_assets_last_season[i][1]-total_assets_last_last_season[i][1])/total_assets_last_last_season[i][1]

        # Mretwd  股票第t月的市场回报
        one_stock['Mretwd'] = mretwd_this_month[i][1]           

        Stock_ALL.append(one_stock)


    # print(Stock_ALL[300])


    # 按股票市值（Size）的中位数把全体股票分成小市值（S）和大市值（B）两组

    Stock_S, Stock_B = Divide_Stock(Stock_ALL,'Size')


    # 按账面市值比（BM）的 30% 和 70% 分位点把全体股票分成低（L）中（N）高（H）三组

    # S
    Stock_SL, Stock_SN, Stock_SH = Divide_Stock(Stock_S,'BM')

    # B
    Stock_BL, Stock_BN, Stock_BH = Divide_Stock(Stock_B,'BM')


    # 以营运利润率（OP）代替账面市值比重复上述步骤 把全体股票分成盈利较弱（W）居中（N2）盈利稳健（R）三组

    # S
    Stock_SW, Stock_SN2, Stock_SR = Divide_Stock(Stock_S,'OP')

    # B
    Stock_BW, Stock_BN2, Stock_BR = Divide_Stock(Stock_B,'OP')


    # 以投资风格（Inv）代替账面市值比重复上述步骤 把全体股票分成保守（C）居中（N3）激进（A）三组

    # S
    Stock_SC, Stock_SN3, Stock_SA = Divide_Stock(Stock_S,'Inv')

    # B
    Stock_BC, Stock_BN3, Stock_BA = Divide_Stock(Stock_B,'Inv')

    SMB_BM = (Mretwd_p(Stock_SH)+Mretwd_p(Stock_SN)+Mretwd_p(Stock_SL))/3 - (Mretwd_p(Stock_BH)+Mretwd_p(Stock_BN)+Mretwd_p(Stock_BL))/3
    SMB_OP = (Mretwd_p(Stock_SR)+Mretwd_p(Stock_SN2)+Mretwd_p(Stock_SW))/3 - (Mretwd_p(Stock_BR)+Mretwd_p(Stock_BN2)+Mretwd_p(Stock_BW))/3
    SMB_Inv = (Mretwd_p(Stock_SC)+Mretwd_p(Stock_SN3)+Mretwd_p(Stock_SA))/3 - (Mretwd_p(Stock_BC)+Mretwd_p(Stock_BN3)+Mretwd_p(Stock_BA))/3
    SMB = (SMB_BM+SMB_OP+SMB_Inv)/3

    HML = (Mretwd_p(Stock_SH)+Mretwd_p(Stock_BH))/2-(Mretwd_p(Stock_SL)+Mretwd_p(Stock_BL))/2
    RMW = (Mretwd_p(Stock_SR)+Mretwd_p(Stock_BR))/2-(Mretwd_p(Stock_SW)+Mretwd_p(Stock_BW))/2
    CMA = (Mretwd_p(Stock_SC)+Mretwd_p(Stock_BC))/2-(Mretwd_p(Stock_SA)+Mretwd_p(Stock_BA))/2

    # print("SMB = "+str(SMB))
    # print("HML = "+str(HML))
    # print("RMW = "+str(RMW))
    # print("CMA = "+str(CMA))


    # print("\n")
    # print(Mretwd_p(Stock_ALL))

    result_tmp = []
    result_tmp.clear()
    result_tmp.append(period)
    result_tmp.append(SMB)
    result_tmp.append(HML)
    result_tmp.append(RMW)
    result_tmp.append(CMA)
    results.append(result_tmp)



    # 将每个周期的 投资组合 分组 及相关指标写入EXCEL 在"分组"中

    write_Excel(Stock_ALL,"Stock_ALL",period)

    write_Excel(Stock_S,"Stock_S",period)
    write_Excel(Stock_B,"Stock_B",period)

    write_Excel(Stock_SL,"Stock_SL",period)
    write_Excel(Stock_SN,"Stock_SN1",period)
    write_Excel(Stock_SH,"Stock_SH",period)

    write_Excel(Stock_BL,"Stock_BL",period)
    write_Excel(Stock_BN,"Stock_BN1",period)
    write_Excel(Stock_BH,"Stock_BH",period)

    write_Excel(Stock_SW,"Stock_SW",period)
    write_Excel(Stock_SN2,"Stock_SN2",period)
    write_Excel(Stock_SR,"Stock_SR",period)

    write_Excel(Stock_BW,"Stock_BW",period)
    write_Excel(Stock_BN2,"Stock_BN2",period)
    write_Excel(Stock_BR,"Stock_BR",period)

    write_Excel(Stock_SC,"Stock_SC",period)
    write_Excel(Stock_SN3,"Stock_SN3",period)
    write_Excel(Stock_SA,"Stock_SA",period)

    write_Excel(Stock_BC,"Stock_BC",period)
    write_Excel(Stock_BN3,"Stock_BN3",period)
    write_Excel(Stock_BA,"Stock_BA",period)

print("\nFinished  \n ")

# 将所有计算所得因子写入EXCEL 在"结果"中

write_Excel_f(results)








############################################# 线性回归部分 #############################################



"""
Created on Sun Nov 29 02:37:40 2020

@author: gy
"""






# regression
# 利用 ./分组 与 ./市场回报数据/rf 与./结果/results.xlsx 中的数据进行处理 得到data.xlsx 其中包含用于计算的五因子与因变量


filePath = '分组'
res = pd.read_excel('结果/results.xls')
print("\n")
print(res)


print("\n \n \n Start to read group data \n")

RfData = pd.read_excel('市场回报数据/rf/TRD_Nrrate.xlsx')
res['Rf'] =RfData['Nrrmtdt']

RMData = pd.read_excel('市场回报数据/rf/TRD_Cnmont.xlsx')
res['RM'] = RMData[RMData['Markettype'] == 5]['Cmretwdos']

for dirpath, dirnames, filenames in os.walk(filePath):
    for file in filenames:
        res[file[:-4]] = 0

for dirpath, dirnames, filenames in os.walk(filePath):
    if dirpath !='分组':
        for file in filenames:
            path = dirpath+'/'+file
            t = str(path)
            t = t[3:]
            print(t)
            #print(file)
            data = pd.read_excel(path)
            if data['Size'].sum()!=0:
                data['sum'] = data['Mretwd'].multiply(data['Size'], axis=0)
                sumData = (data['sum'].sum()) / (data['Size'].sum())
                # print(sumData)
            else:
                sumData = 0

            res[file[:-4]] = res.apply(lambda x: sumData if x['period'] == dirpath[3:] else x[file[:-4]],axis = 1)

    res.to_excel('结果/data.xlsx')


print("\nFinished  \n ")


print("\nStart regression \n")


# redundant
# 得到五因子各因子与其他四因子回归的结果


data1 = pd.read_excel('结果/data.xlsx',sheet_name='Sheet1')
data1['RMMinusRf'] = data1['RM'] - data1['Rf']
# lm = ols('RiMinusRf~RMMinusRf+SMB+HML+CMA+RMW', data=data1).fit()
lm1 = ols('RMMinusRf~SMB+CMA+RMW+HML', data=data1).fit()
lm2 = ols('SMB~RMMinusRf+CMA+RMW+HML', data=data1).fit()
lm3 = ols('CMA~SMB+RMMinusRf+RMW+HML', data=data1).fit()
lm4 = ols('RMW~SMB+CMA+RMMinusRf+HML', data=data1).fit()
lm5 = ols('HML~SMB+CMA+RMW+RMMinusRf', data=data1).fit()
print(lm1.summary())
print(lm2.summary())
print(lm3.summary())
print(lm4.summary())
print(lm5.summary())



# r
# 得到回归结果

data2 = pd.read_excel('结果/data.xlsx',sheet_name='Sheet1')
data2.drop('period', axis=1, inplace=True)
data2.drop('Unnamed: 0', axis=1, inplace=True)


for i in data2:
    if i != 'Rf' and i != 'RM':
        print('\n******************************** '+i+' ************************************\n')
        data2['RiMinusRf'] = data2[i] - data2['Rf']
        data2['RMMinusRf'] = data2['RM'] - data2['Rf']
        lm = ols('RiMinusRf~RMMinusRf+SMB+HML+CMA+RMW', data=data2).fit()
        print(lm.summary())

print("\nFinished  \n ")