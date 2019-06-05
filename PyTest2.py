
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

URL_ = 'http://stats.espncricinfo.com/ci/engine/player/253802.html?class=2;template=results;type=batting'

#from stats2 to stats 1
df_list_ = pd__.read_html(URL_)
df_1_ = DF(df_list_[2])
df_1_ = df_1_.set_index('Unnamed: 0')
df_2_ = DF(df_list_[3])
df_2_ = df_2_.set_index('Grouping')
df_bat_1st_ = df_2_.loc[['matches batting first']]
df_1_ = df_1_.append(df_bat_1st_)
val_ = df_bat_1st_

print(df_1_)
# print(val_)

