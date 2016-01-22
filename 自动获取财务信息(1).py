# python操作Excel word
# 分类： Utility2012-10-08 21:42 554人阅读 评论(0) 收藏 举报
# excelpythondeletetuples文档insert
# 跟网上其他的程序是一样的，这里分别增加了一个全局替换的函数。
# 可以进行整个文档的替换，这样可以实现翻译，查错等全局替换的操作。


import os,types,pickle
from win32com.client import Dispatch
import win32com.client
import urllib.request
from bs4 import BeautifulSoup
from collections import defaultdict

DEBUG=True

def get_stock_info(stock_code):
#判断是深圳，香港或者上海的股票
    try:
        if len(stock_code)==5:
            stock_code='hk'+stock_code
        elif stock_code.startswith('6'):
            stock_code='sh'+stock_code
        else:
            stock_code='sz'+stock_code
        url = "http://hq.sinajs.cn/list=%s" % stock_code
        response= urllib.request.urlopen(url)
        data=response.readline().decode('gb2312')
        response.close()
        data=data.split('=')[1]
        data_list=data.split(',')
        res=dict()
        if stock_code.startswith('hk'):
            res['date']=data_list[-2]
            res['price']=data_list[6]
            res['name']='hk'+data_list[1]
            print(res['name'],res['price'])
        else:
            res['date']=data_list[-3]
            res['price']=data_list[3]
            res['name']=data_list[0][1:]
            print(res['name'],res['price'])
        #print(data_list)
        
        return res
    except Exception as e:
        print(">>>>>> Exception: " + str(e))



class Excel:
    """A utility to make it easier to get at Excel."""

    def __init__(self, filename=None):
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        self.xlApp.Visible = False
        self.xlApp.DisplayAlerts = False  #搜索不到时不提示对话框
        
        if filename:
          self.filename = filename
          self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
          self.xlBook = self.xlApp.Workbooks.Add()
          self.filename = '' 

    def save(self, newfilename=None):
      if newfilename:
          self.filename = newfilename
          self.xlBook.SaveAs(newfilename)
      else:
          self.xlBook.Save()
          
    def close(self):
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def getCell(self, sheet, row, col):
        "Get value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value


    def setCell(self, sheet, row, col, value):
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value


    def getRange(self, sheet, row1, col1, row2, col2):
        "return a 2d array (i.e. tuple of tuples)"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value

    def addPicture(self, sheet, pictureName, Left, Top, Width, Height):
        "Insert a picture in sheet"
        sht = self.xlBook.Worksheets(sheet)
        sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)

    def cpSheet(self,source,name):
        '''复制source sheet到name sheet'''
        shts = self.xlBook.Worksheets
        try: 
            shts(name)
        except Exception: #没有名字为name的 sheet才复制
            shts(source).Copy(Before=shts(source)) #复制source sheet并放到before前
            ws=shts(source+' (2)')
            ws.Name=name
    
    def getUsedRange(self, sheetindex):
        sht = self.xlBook.Worksheets(sheetindex)  
        return sht.UsedRange.Rows.Count,sht.UsedRange.Columns.Count

    def getSheetCount(self):
        return self.xlBook.sheets.Count

    def replace(self, sheetindex, oldStr, newStr):
        """Find the oldStr and replace with the newStr.
        """
        #Replace(oldStr, newStr, LookAt, SearchOrder, MatchCase, MatchByte, SearchFormat, ReplaceFormat)
        # print self.xlBook.Worksheets(sheetindex).Name
        #self.xlBook.Worksheets(sheetindex).Activate   #activate the current sheet
        #self.xlApp.Selection.Replace(oldStr, newStr)   #replace only activate sheet

        sht = self.xlBook.Worksheets(sheetindex)
        sht.Cells.Replace(oldStr, newStr) #replace every sheet

    def find_label(self,sheet,label):
        '''找到数据项目在第一列的多少行
        '''
        for i in range(1,50):
            if self.getCell(sheet,i,1)==label:
                return i


def full_stock_code(code):
    if len(code)==5:
        return 'hk'+code
    elif code.startswith('6'):
        return 'sh'+code
    else:
        return 'sz'+code

def get_soup(url):
    response= urllib.request.urlopen(url)
    html=response.read().decode('utf8')
    return BeautifulSoup(html,"html.parser")

def get_gu_ben(code,excel,sheet,years,parameters):
    url = 'http://f10.eastmoney.com/f10_v2/CapitalStockStructure.aspx?code=%s' % code
    soup = get_soup(url)
    fd=soup.find(id='lngbbd_Table').contents #股本的table
    
    zong_gu_bens=fd[1].contents #总股本
    dates=fd[0].contents #日期
    liu_tong_gus=fd[7].contents #流通股本
    reasons=fd[9].contents #原因
    year_reason=defaultdict(lambda:'')
    
    for i,child in enumerate(dates): #日期比其他行多了一个第二列
        try:
            reason=reasons[i-1].string
            year=child.string[2:4]
            if reason!='定期报告' and year_reason[year].find(reason)==-1:
                #避免添加重复原因
                year_reason[year]+=reason.replace('上市','')+','

            if  child.string[-5:]=='12-31' : #年末日期
                col=years[int(year)]
                
                zong_gu_ben=zong_gu_bens[i-1].string.replace(',','')
                excel.setCell(sheet,parameters['总股本'],col,round(float(zong_gu_ben)/10000,1))

                liu_tong_gu=liu_tong_gus[i-1].string.replace(',','')
                excel.setCell(sheet,parameters['流通股'],col,round(float(liu_tong_gu)/10000,1))

                #print(child.string,liu_tong_gu)
        except Exception:
            pass

    # 第一个总是写入最新的总股本
    zong_gu_ben=zong_gu_bens[1].string.replace(',','')
    excel.setCell(sheet,parameters['总股本'],2,round(float(zong_gu_ben)/10000,1))

    liu_tong_gu=liu_tong_gus[1].string.replace(',','')
    excel.setCell(sheet,parameters['流通股'],2,round(float(liu_tong_gu)/10000,1))
    for key in year_reason:
        try:
            excel.setCell(sheet,parameters['变动原因'],years[int(key)],year_reason[key])
        except ValueError:
            pass
    #if DEBUG:
        #print(liu_tong_gus)
        #print(dates)
        #Sprint(year_reason)


def get_finance(code,excel,sheet,years,parameters):
    url='http://f10.eastmoney.com/f10_v2/FinanceAnalysis.aspx?code=%s' % code
    soup = get_soup(url)
    main_table=soup.find(id='F10MainTargetDiv').contents[3].contents
    dates=main_table[0]
    print(dates)

def analyze_template(excel,sheet):
    years=dict() #key 年，value 位于哪一列
    parameters=dict() #key 参数名字，value 位于哪一行

    for i in range(2,13):
        years[int(excel.getCell(sheet,1,i))]=i
    

    for i in range(2,30):
        parameters[excel.getCell(sheet,i,1)]=i
    # if DEBUG:
    #     print(parameters)
    #     print(years)

    return years,parameters


def get_info_from_csv(code,excel,sheet,years,parameters):
    current_folder=os.path.dirname(os.path.abspath(__file__))
    path=os.path.join(current_folder,'财务指标'+code+'.csv')
    csv_file=Excel(path)
    csv_parameters=dict()
    for i in range(2,21):
        print(csv_file.getCell('财务指标'+code,i,1))
        csv_parameters[csv_file.getCell('财务指标'+code,i,1)]=i
    print(csv_parameters)

    

    csv_file.close()
    # with open(path) as f:
    #     f_csv = csv.reader(f)
    #     headings = next(f_csv)
    #     print(headings)
        # Row = namedtuple('Row', headings)
        # for r in f_csv:
        #     row = Row(*r)
        #     print(row)

if __name__ == "__main__":
    file_name='机场航运.xlsx'
    stock_code='600004'
    #stock_name='天地科技'
    stock_name=get_stock_info(stock_code)['name']
    print(stock_name)

    #复制模版
    current_folder=os.path.dirname(os.path.abspath(__file__))
    file_path=os.path.join(current_folder,file_name)
    excel_file=Excel(file_path)
    excel_file.cpSheet('模版',stock_name)
    
    years,parameters=analyze_template(excel_file,stock_name)
    #get_gu_ben(full_stock_code(stock_code),excel_file,stock_name,years,parameters)
    get_info_from_csv(stock_code,excel_file,stock_name,years,parameters)
    
    excel_file.save()
    excel_file.close()
    print('完成')
