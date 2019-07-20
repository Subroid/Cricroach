
class TopPerformancesFetcher :

    import urllib.request as Ureq
    import pandas as Pnds
    from pandas import DataFrame as DFrame
    from openpyxl import Workbook as Wb
    from openpyxl.utils import dataframe as Odf
    from bs4 import BeautifulSoup as Bsoup
    from bs4 import Tag
    import re as Re
    import openpyxl.styles as Ostyles

    def __init__(self):
        self

    def get_head_to_head_last_5_matches_links(self, match_url):
        url_resp= self.Ureq.urlopen(match_url)
        PARSER= 'html.parser'
        bsoup= self.Bsoup(url_resp, PARSER)
        list_links= list()
        span_tags_list= bsoup.find_all('div', attrs={'class': 'sub-module last-games head-to-head cricket'})
        span_tag= span_tags_list[0]
        table_list= span_tag.find_all('table')
        table_tag= table_list[0]
        td_tags_list= table_tag.findAll('td', attrs={'class': 'outcome'})
        for i in range(len(td_tags_list)):
            td_tag= td_tags_list[i]
            match_result_str = td_tag.string
            if ('lost' in str(match_result_str).lower()
                    or 'beat' in str(match_result_str).lower()):
                a_tag= td_tag.find('a')
                link = a_tag['href']
                list_links.append(link)
            else:
                continue
        print(len(list_links))
        return list_links