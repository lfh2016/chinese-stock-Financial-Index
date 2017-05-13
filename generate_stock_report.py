import os
import urllib.request

import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from pandas import ExcelWriter

current_folder = os.path.dirname(os.path.abspath(__file__))


# 拟分配的利润或股利，可以得到分红的数额

class Stock():
    caiwu_folder = os.path.join(current_folder, '财务数据')
    need_items = ['营业总收入(万元)', '研发费用(万元)', '财务费用(万元)', '净利润(万元)_y',
                  '归属于母公司所有者的净利润(万元)', '总资产(万元)', '总负债(万元)', '流动资产(万元)', '流动负债(万元)'
        , '股东权益不含少数股东权益(万元)', '净资产收益率加权(%)', ' 支付给职工以及为职工支付的现金(万元)',
                  '经营活动产生的现金流量净额(万元)', ' 投资活动产生的现金流量净额(万元)', '应收账款(万元)'
        , '存货(万元)', '开发支出(万元)', '归属于母公司股东权益合计(万元)', '所有者权益(或股东权益)合计(万元)'
        , '投资收益(万元)_x', '实收资本(或股本)(万元)', '每股净资产(元)']
    # ,'支付给职工以及为职工支付的现金(万元)'

    items_new_name = ['营业收入', '研发费用', '财务费用', '净利润', '归属净利润', '总资产', '总负债', '流动资产',
                      '流动负债', '股东权益', 'ROE', ' 职工薪酬', '经营现金流', '投资现金流'
        , '应收账款', '存货', '开发支出', '归属股东权益', '权益合计'
        , '投资收益', '总股本', '每股净资产']  # 名字和上面列表一一对应

    report_end_year = 2016

    def __init__(self, code, name):
        self.code = code
        self.name = name
        # self.dframe=dframe
        self.cwzb_path = os.path.join(self.caiwu_folder, '财务指标' + self.code + '.csv')
        self.zcfzb_path = os.path.join(self.caiwu_folder, '资产负债表' + self.code + '.csv')
        self.lrb_path = os.path.join(self.caiwu_folder, '利润表' + self.code + '.csv')
        self.xjllb_path = os.path.join(self.caiwu_folder, '现金流量表' + self.code + '.csv')
        self.cwzb_url = 'http://quotes.money.163.com/service/zycwzb_%s.html?type=report' % self.code  # 财务总表
        self.zcfzb_url = 'http://quotes.money.163.com/service/zcfzb_%s.html' % self.code  # 资产负债表
        self.lrb_url = 'http://quotes.money.163.com/service/lrb_%s.html' % self.code  # 利润表
        self.xjll_url = 'http://quotes.money.163.com/service/xjllb_%s.html' % self.code  # 现金流量表
        self.url_reports = {self.cwzb_url: self.cwzb_path, self.zcfzb_url: self.zcfzb_path,
                            self.lrb_url: self.lrb_path, self.xjll_url: self.xjllb_path}
        if self.code.startswith('6'):
            stock_code = 'sh' + self.code
        else:
            stock_code = 'sz' + self.code
        self.fh_url = 'http://f10.eastmoney.com/f10_v2/BonusFinancing.aspx?code=%s' % stock_code
        # url和对应的文件路径

    def save_xls(self, dframe):  # 把数据写到已行业命名的excel文件的名字sheet
        xls_path = os.path.join(current_folder, self.name + '.xlsx')
        if os.path.exists(xls_path):  # excel 文件已经存在
            book = load_workbook(xls_path)
            writer = pd.ExcelWriter(xls_path, engine='openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            dframe.to_excel(writer, self.name)
            writer.save()
        else:  # 文件还不存在
            writer = ExcelWriter(xls_path)
            dframe.to_excel(writer, self.name)
            writer.save()

    def download_if_need(self, url, path):
        if not os.path.exists(path):
            urllib.request.urlretrieve(url, path)

    def doanload_stock_info(self):
        for url in self.url_reports:
            self.download_if_need(url, self.url_reports[url])

    def _generate_report(self):
        data_frames = []
        for key in self.url_reports:
            data_frames.append(pd.read_csv(self.url_reports[key], encoding="gbk", index_col=0).T)
        merge_frame = None
        for frame in data_frames:
            if merge_frame is None:
                merge_frame = frame
            else:
                merge_frame = pd.merge(merge_frame, frame, left_index=True, right_index=True)
        merge_frame = merge_frame[self.need_items]  # 只取需要的指标
        merge_frame.columns = self.items_new_name  # 重命名列
        for index, row in merge_frame.iterrows():
            try:
                merge_frame.loc[index, 'ROE'] = 10000 * float(row['ROE'])
            except Exception:
                pass
        try:
            merge_frame['自由现金流'] = pd.to_numeric(merge_frame['经营现金流']) + pd.to_numeric(merge_frame['投资现金流'])
        except Exception as e:
            merge_frame['自由现金流'] = 0
            print(self.name + '自由现金流计算失败')

        try:
            merge_frame['负债率'] = pd.to_numeric(merge_frame['总负债']) / pd.to_numeric(merge_frame['总资产'])
            merge_frame['负债率'] = merge_frame['负债率'].round(2)
        except Exception as e:
            merge_frame['负债率'] = 0.66
            print(self.name + '负债率计算失败')
        try:
            merge_frame['流动比率'] = pd.to_numeric(merge_frame['流动资产']) / pd.to_numeric(merge_frame['流动负债'])
            merge_frame['流动比率'] = merge_frame['流动比率'].round(1)
        except Exception as e:
            merge_frame['流动比率'] = 0.66
            print(self.name + '流动比率计算失败')
        merge_frame = merge_frame[['营业收入', '流动比率', '负债率', '经营现金流', '投资现金流', '自由现金流',
                                   '净利润', '投资收益', ' 职工薪酬', '财务费用',
                                   '研发费用', '开发支出', 'ROE', '总股本', '每股净资产']]
        merge_frame = merge_frame.T
        years = []
        for y in range(self.report_end_year, 2005, -1):
            date = str(y) + '-12-31'
            if date in merge_frame.columns:  # 如果存在该时间的数据
                years.append(date)
        merge_frame = merge_frame[years]  # 只取需要的时间
        merge_frame = merge_frame.applymap(self.convert2yi)
        self.save_xls(merge_frame)

    def generate_report(self):
        self.doanload_stock_info()
        self._generate_report()


    @staticmethod
    def convert2yi(value):
        try:
            if float(value) > 10 or float(value) < -10:
                return round(float(value) / 10000, 1)
            else:
                return value
        except Exception as e:
            return value

    def get_soup(self):
        response = urllib.request.urlopen(self.fh_url)
        html = response.read().decode('utf8')
        return BeautifulSoup(html, "html.parser")

    def get_3year_average_fh(self):  # 单位为万元
        try:
            print('获取%s的分红信息' % self.name)
            soup = self.get_soup()
            fd = soup.find(id='lnfhrz')
            fd = fd.next_sibling.next_sibling.contents[1].contents
            sum_fh = float(fd[1].contents[1].string.replace(',', '')) + float(
                fd[2].contents[1].string.replace(',', '')) + float(fd[3].contents[1].string.replace(',', ''))
            # print(sum_fh)
            return round(sum_fh / 3)
        except Exception as e:  # 没有分红
            return 0


def update_fhlv():
        stocks_path = os.path.join(current_folder, '筛选后股票的财务报表', '筛选后的股票列表.xlsx')
        stocks = pd.read_excel(stocks_path, index_col=0)
        stocks['3年平均分红'] = 0
        for index, row in stocks.iterrows():
            s = Stock('%06d' % index, row['名字'])
            stocks.loc[index, '3年平均分红'] = s.get_3year_average_fh()
        stocks['平均分红率'] = stocks['3年平均分红'] * 100 / (stocks['总股本(万)'] * stocks['价格'])
        stocks['平均分红率'] = stocks['平均分红率'].round()
        stocks.to_excel(stocks_path)


def generate_reports():  # 根据股票列表,生成报表
    stocks_path = os.path.join(current_folder, '筛选后股票的财务报表', '筛选后的股票列表.xlsx')
    stocks = pd.read_excel(stocks_path, index_col=0)
    for index, row in stocks.iterrows():
        s = Stock('%06d' % index, row['名字'])
        print('正在生成' + row['名字'] + '的报表')
        s.generate_report()

if __name__ == '__main__':
    # s=Stock(sys.argv[1],sys.argv[2])  #股票代码，名字
    s = Stock('600660', '福耀玻璃')
    # # #s.doanload_stock_info()
    s.generate_report()
    # # s=Stock('000568','泸州老窖','白酒')
    # print(s.get_3year_average_fh())
    # generate_reports()
    # update_fhlv()
    print('完成')
    # 测试哦
