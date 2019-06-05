


import urllib.request as UReq
from bs4 import BeautifulSoup as BSoup
import pandas as pd__
from openpyxl import Workbook as wb__
import openpyxl.utils.dataframe as odf__
from openpyxl import load_workbook
import re as re__
from bs4 import Tag as TAG
from pandas import DataFrame as DF
from openpyxl.worksheet.worksheet import Worksheet as WS

URL_ = 'https://kite.zerodha.com/dashboard'
url_resp_ = UReq.urlopen(URL_)
bsoup_ = BSoup(url_resp_, 'html.parser')
val_ =  bsoup_.prettify()
print(val_)