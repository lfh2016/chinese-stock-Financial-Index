# chinese-stock-Financial-Index
计算中国A股所有股票的3年平均财务指标,并可以按照3年平均市盈率过滤股票

使用前提:
1. 下载并安装python3版本的Anaconda（https://www.anaconda.com/download/）
2. 在当前目录打开命令行，输入 pip install -r requirement。


使用方法
1. 解压'finance2017.7z'文件到当前文件夹，解压后为finance2017文件夹，包含所有a股公司到2017年的财务数据
2. 运行calcu_3year_average_pe.py 通过3年平均收益率筛选股票，默认为市盈率范围为2-20，可以修改。
生成类似 '2017-06-09-3年平均市盈率在2和20之间的公司.xlsx'文件名的文件，里面包含筛选出来的公司。
直接用excel打开即可查看