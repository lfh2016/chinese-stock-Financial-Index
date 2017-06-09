import pandas as pd

# 比较2次筛选的股票池，得到新入池的公司
old_day = '2016-07-12'
new_day = '2016-10-08'


def new_company_by_compare(new_day, old_day):
    old_pool_path = '3年平均市盈率在2和20之间的公司%s.xlsx' % (old_day)
    new_pool_path = '3年平均市盈率在2和20之间的公司%s.xlsx' % (new_day)
    old_pool = pd.read_excel(old_pool_path)
    new_pool = pd.read_excel(new_pool_path)
    new_index = new_pool.index.difference(old_pool.index)
    new_company = new_pool.reindex(new_index)
    new_company_file = '%s比%s新入池的公司.xlsx' % (new_day, old_day)
    new_company.to_excel(new_company_file)


if __name__ == '__main__':
    new_company_by_compare(new_day, old_day)
