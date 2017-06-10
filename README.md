# chinese-stock-Financial-Index
计算中国A股所有股票的3年平均财务指标,并可以按照3年平均市盈率过滤股票

使用前提:
1. 安装Python3
2. 安装pandas
3. lxml也是必须的，正常情况下安装了Anaconda后无须单独安装，如果没有可执行：pip install lxml
4. tushare

建议安装Anaconda（http://www.continuum.io/downloads），一次安装包括了Python环境和全部依赖包，减少问题出现的几率。

**2017-06-09 pip tushare最新版本（0.7.9）的获取今日股票价格函数有问题，所以直接把它github的源代码拷过来（这个上面的没有这个问题）。**

使用方法
1. 解压'finance2016.7z'文件到当前文件夹，解压后为finance2016文件夹，包含所有a股公司到2016年的财务数据
2. 运行calcu_3year_average_pe.py 通过3年平均收益率筛选股票，默认为市盈率范围为2-20，可以修改。
生成类似 '2017-06-09-3年平均市盈率在2和20之间的公司.xlsx'文件名的文件，里面包含筛选出来的公司。
直接用excel打开即可查看