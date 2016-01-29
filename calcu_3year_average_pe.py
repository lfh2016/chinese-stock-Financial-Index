import os
import urllib.request
from datetime import datetime

import pandas as pd
import tushare as ts

current_folder=os.path.dirname(os.path.abspath(__file__))
caiwu_folder=os.path.join(current_folder,'财务数据')
calcu_average_profit_end_year= 2014 #计算3年平均利润的截止年



def create_folder_if_need(path):
    if not os.path.exists(path):  # 如果该文件夹不存在，创建文件夹
        os.makedirs(path)
    elif not os.path.isdir(path):
        os.makedirs(path)


def download_if_need(code,url,name):
    path=os.path.join(caiwu_folder,name+code+'.csv')
    if not os.path.exists(path):
        urllib.request.urlretrieve(url, path)


def download_stock_data(code):
    cwzb_url='http://quotes.money.163.com/service/zycwzb_%s.html?type=report'%code
    #财务指标
    create_folder_if_need(caiwu_folder)
    download_if_need(code,cwzb_url,'财务指标')


def calcu_3year_average_profit(code,year):
    download_stock_data(code)
    data=pd.read_csv(os.path.join(caiwu_folder,'财务指标'+code+'.csv'),encoding = "gbk",index_col=0)
    data=data.T
    average_profit=0
    for i in range(year-2,year+1):
        average_profit+=float(data['净利润(万元)'][str(i)+'-12-31'])
    average_profit /=3
    print(average_profit)
    return average_profit


def calcu_all_stocks_3year_average_profit(year): #生成3年平均利润列表
    path=os.path.join(current_folder,'股票列表.csv')
    if not os.path.exists(path):
        data=ts.get_stock_basics()
        lie=['名字','行业','地区','市盈率','流通股本','总股本(万)',
        '总资产(万)','流动资产','固定资产','公积金','每股公积金','每股收益','每股净资','市净率','上市日期']
        data.columns=lie
        data.index.names=['代码']
        data.to_csv(path)
        
    data=pd.read_csv(path,encoding = "gbk",index_col=0)
    print(data)
    data['平均利润']=0
    for index, row in data.iterrows():
        data.loc[index,'平均利润']=calcu_3year_average_profit('%06d'%index,year)

        print('完成%s'%index)
    data.to_csv(os.path.join(current_folder,'3年平均利润.csv'))


def filter_stock_by_average_pe(min, max):
    path=os.path.join(current_folder,'3年平均利润.csv')
    if not os.path.exists(path): #没有就生成3年平均利润列表
        calcu_all_stocks_3year_average_profit(calcu_average_profit_end_year)

    gplb=pd.read_csv(path,encoding = "gbk",index_col=0)

    now=datetime.now()
    today=str(now)[:11]
    current_sec=str(now)[:18].replace('-','_').replace(':','_')
    price_path=os.path.join(current_folder,today+'股票价格.csv')
    if not os.path.exists(price_path):
        ts.get_today_all().set_index('code').to_csv(price_path)

    current_price=pd.read_csv(price_path,encoding = "gbk",index_col=0)
    current_price= current_price[['trade']]
    current_price.columns=['价格']
    gplb=gplb[['名字','行业','地区','流通股本','总股本(万)','总资产(万)','流动资产','固定资产','每股净资','市净率','上市日期','平均利润']]
    data=pd.merge(gplb,current_price,left_index=True, right_index=True)
    data['平均市盈率']=data['总股本(万)']*data['价格']/data['平均利润']
    data = data[data['平均市盈率'] < max]
    data = data[data['平均市盈率'] > min]
    data['平均市盈率']=data['平均市盈率'].round()
    data['平均利润']=data['平均利润'].round()
    data['市净率']=data['市净率'].round(1)
    data['固定资产']=data['固定资产'].round()
    data['流动资产']=data['流动资产'].round()
    data['总股本(万)']=data['总股本(万)'].round()
    data['流通股本']=data['流通股本'].round()
    average_pe_file = os.path.join(current_folder, today + '3年平均市盈率在%s和%s之间的公司.xlsx' % (min, max))
    data.to_excel(average_pe_file)


filter_stock_by_average_pe(2, 20)
print('完成')