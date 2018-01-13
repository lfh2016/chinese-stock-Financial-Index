import os
import urllib.request
from datetime import datetime
from time import sleep

import pandas as pd

import tushare as ts

current_folder = os.path.dirname(os.path.abspath(__file__))

calcu_average_profit_end_year = 2016  # 计算平均利润的截止年,包括该年
caiwu_folder = os.path.join(current_folder, 'finance%s' % calcu_average_profit_end_year)
DEBUG = True

today = str(datetime.now())[:10]


def create_folder_if_need(path):
    if not os.path.exists(path):  # 如果该文件夹不存在，创建文件夹
        os.makedirs(path)
    elif not os.path.isdir(path):
        os.makedirs(path)


def download_if_need(code, url):
    path = os.path.join(caiwu_folder, code + '.csv')
    if not os.path.exists(path):
        urllib.request.urlretrieve(url, path)
        sleep(1.5)  # 间隔一段时间，防止服务器关闭连接


def download_stock_data(code):
    cwzb_url = 'http://quotes.money.163.com/service/zycwzb_%s.html?type=report' % code
    # 财务指标
    create_folder_if_need(caiwu_folder)
    download_if_need(code, cwzb_url)


def calcu_3year_average_profit(code, year):
    download_stock_data(code)
    data = pd.read_csv(os.path.join(caiwu_folder, code + '.csv'), encoding="gbk",
                       index_col=0)
    data = data.T
    average_profit = 0
    for i in range(year - 2, year + 1):
        # 确认过，这里的'净利润(万元)'就是归属净利润
        average_profit += float(data['净利润(万元)'][str(i) + '-12-31'])
    average_profit /= 3
    # print(average_profit)
    return average_profit


def calcu_all_stocks_3year_average_profit(year):  # 生成3年平均利润列表
    path = os.path.join(current_folder, '股票列表%s.csv' % today)
    if not os.path.exists(path):
        data = ts.get_stock_basics()
        lie = ['名字', '行业', '地区', '市盈率', '流通股本', '总股本',
               '总资产(万)', '流动资产', '固定资产', '公积金', '每股公积金', '每股收益', '每股净资',
               '市净率', '上市日期', '未分利润', '每股未分配', '收入同比(%)', '利润同比(%)',
               '毛利率(%)', '净利润率(%)', '股东人数']
        data.columns = lie
        data.index.names = ['代码']
        data.to_csv(path, encoding='utf-8')

    data = pd.read_csv(path, encoding='utf-8', index_col=0)
    # print(data)
    data['平均利润'] = 0
    for index, row in data.iterrows():
        try:
            data.loc[index, '平均利润'] = calcu_3year_average_profit('%06d' % index, year)
        except Exception as e:
            print(e)
            data.loc[index, '平均利润'] = 0

        print('完成%s' % index)
    data.to_csv(os.path.join(current_folder, '3年平均利润及其他财务指标%s.csv' % today), encoding='utf-8')


def filter_stock_by_average_pe(min, max):
    path = os.path.join(current_folder, '3年平均利润及其他财务指标%s.csv' % today)
    if not os.path.exists(path):  # 没有就生成3年平均利润列表
        calcu_all_stocks_3year_average_profit(calcu_average_profit_end_year)

    gplb = pd.read_csv(path, index_col=0, encoding='utf-8')

    # 获取当前股票价格
    price_path = os.path.join(current_folder, today + '股票价格.csv')
    if not os.path.exists(price_path):
        ts.get_today_all().set_index('code').to_csv(price_path, encoding="utf-8")

    current_price = pd.read_csv(price_path, encoding="utf-8", index_col=0)
    current_price = current_price[['trade']]
    current_price.columns = ['价格']
    gplb = gplb[
        ['名字', '行业', '地区', '流通股本', '总股本', '总资产(万)', '流动资产', '固定资产', '每股净资', '市净率', '上市日期',
         '平均利润']]

    data = pd.merge(gplb, current_price, left_index=True, right_index=True)
    # 因为这里的平均利润单位是万元，而总股本单位是亿，价格单位是元
    data['平均市盈率'] = data['总股本'] * data['价格'] * 10000 / data['平均利润']
    print('\n%s:' % today)
    print()
    print('%d个公司' % data.shape[0])
    print('3年市盈率中位数%.1f' % round(data['平均市盈率'].median(), 1))
    print('市净率中位数%.1f' % round(data['市净率'].median(), 1))
    data = data[data['平均市盈率'] < max]
    data = data[data['平均市盈率'] > min]
    data['平均市盈率'] = data['平均市盈率'].round(1)
    data['平均利润'] = data['平均利润'].round()
    data['市净率'] = data['市净率'].round(1)
    data['固定资产'] = data['固定资产'].round()
    data['流动资产'] = data['流动资产'].round()
    data['总股本'] = data['总股本'].round()
    data['流通股本'] = data['流通股本'].round()
    average_pe_file = os.path.join(current_folder, today + '-3年平均市盈率在%s和%s之间的公司.xlsx' % (min, max))
    data.to_excel(average_pe_file)


if __name__ == '__main__':
    filter_stock_by_average_pe(1, 20)  # 这个函数是根据平均pe过滤股票
