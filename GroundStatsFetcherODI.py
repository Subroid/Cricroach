

import urllib.request as url_req__
from bs4 import BeautifulSoup as bsoup__
import pandas as pd__
from openpyxl import Workbook as wb__
import openpyxl.utils.dataframe as odf__
from openpyxl import load_workbook
from pandas import DataFrame as DF
#^TODO deletable
from openpyxl.worksheet.worksheet import Worksheet as WS
#^TODO deletable


url_of_match_ = 'http://www.espncricinfo.com/series/8039/game/1144496/australia-vs-india-14th-match-icc-cricket-world-cup-2019'
url_main_domain_ = 'http://www.espncricinfo.com/'
url_resp_ = url_req__.urlopen(url_of_match_)
PARSER_ = 'html.parser'

EXCEL_FILE_LOC_ = 'E:/D11/D11/'
CURRNET_DATE_ = '08+Jun+2019'
TARGET_DATE_ = '02+Jun+2017'


bsoup_ = bsoup__(url_resp_, PARSER_)

h4_tag_ = bsoup_.find('h4')
a_tag_ = h4_tag_.find('a')
ground_name_ = a_tag_.string
url_ground_ = a_tag_['href']

excel_file_ = EXCEL_FILE_LOC_ + ground_name_ + ' stats.xlsx'

url_resp_ = url_req__.urlopen(url_ground_)
bsoup_ = bsoup__(url_resp_, PARSER_)
div_tag_ = bsoup_.find('div', {'id': 'recs'})
table_tag_ = div_tag_.find('table')
tr_tags_ = table_tag_.find_all('tr', {'class': 'islast1'})
tr_tag_ = tr_tags_[1]
a_tags_ = tr_tag_.find_all('a')
a_tag_ = a_tags_[8]
url_ground_stats_sub_domain_ = a_tag_['href']
url_ground_stats_1_ = url_main_domain_ + url_ground_stats_sub_domain_
url_ground_stats_recent_years_ = url_ground_stats_1_ + ';spanmax1=' + CURRNET_DATE_ + ';spanmin1=' + TARGET_DATE_ + ';spanval1=span;template=results'
'''http://stats.espncricinfo.com/ci/engine/ground/57219.html?class=2;spanmax1=02+Jun+2019;spanmin1=02+Jun+2017;spanval1=span;template=results;type=aggregate'''

df_list_ = pd__.read_html(url_ground_stats_recent_years_)
df_1_ = df_list_[2]
df_2_ = df_list_[3]
df_2_ = df_2_.dropna(axis=1, how='all')
df_2_ = df_2_.fillna("")

wb_ = wb__()
ws_ = wb_['Sheet']
ws_.title = 'Recent Years'
deployable_odf_ = odf__.dataframe_to_rows(df_1_, index=False, header=True)
for df_row in deployable_odf_:
    ws_.append(df_row)
wb_.save(excel_file_)
wb_.close()

wb_ = load_workbook(excel_file_)
ws_ = wb_['Recent Years']
ws_.append([""])
deployable_odf_ = odf__.dataframe_to_rows(df_2_, index=False, header=True)
for df_row in deployable_odf_:
    ws_.append(df_row)
wb_.save(excel_file_)
wb_.close()













