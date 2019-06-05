
import urllib.request as url_req__
from bs4 import BeautifulSoup as bsoup__
import pandas as pd__
from openpyxl import Workbook as wb__
import openpyxl.utils.dataframe as odf__
from openpyxl import load_workbook
import re as re__
from bs4 import Tag as TAG
from pandas import DataFrame as DF
from openpyxl.worksheet.worksheet import Worksheet as WS
from pandas.core.series import Series
import cricpy.analytics as ca

URL_ = 'http://stats.espncricinfo.com/ci/engine/player/253802.html?class=2;template=results;type=batting;view=innings'

''' #todo catching KeyError
error_ = ''
try:
    val_ = df_.loc[['0 runs']]
except KeyError as e:
    error_ =  str(e)
print(len(error_))
'''
# val_ = df_.at['0 runs', 'Dis'] + df_.at['1-9 runs', 'Dis'] + df_.at['10-19 runs', 'Dis']
# df_ = df_.set_index('Grouping')
# df_ = df_.replace(to_replace="-", value="")
# df_ = df_list_[3]
# print(df_.iloc[:, -3:-2])
# df_.iloc[:, -3:-2] = df_.iloc[:, -3:-2].apply(pd__.to_numeric)
# df_.to_excel('E:/D11/testexl.xlsx')


df_list_ = pd__.read_html(URL_)
df_ = DF(df_list_[3])
df_last_ = DF(df_.tail(1))
found_19_ = 0
found_50_ = 0
found_100_ = 0
# val_ = df_last_.at[85, 'Runs']
val_ = len(df_.index)-1

def find_last_match(greater_than, less_than, df_index_size):
    j = 0
    for i in range(df_index_size):
        val = df_.at[val_-i, 'Runs']
        # todo if val contains * then parse and contains DNB etc then pass
        if(str(val).find('*')):
            val = str(val).split('*')[0]
            try:
                val = int(val)
                j = j+1
                if(val>greater_than and val<less_than):
                    print(val)
                    break
            except ValueError as e:
                pass
    return j

found_19_ = find_last_match(-1, 20, val_)
found_50_ = find_last_match(50, 100, val_)
found_100_ = find_last_match(100, 300, val_)
print(found_19_)
print(found_50_)
print(found_100_)
DF
