import urllib.request as __urlreq

_url = __urlreq.urlopen("http://www.espncricinfo.com/series/8039/game/1144483/england-vs-south-africa-1st-match-icc-cricket-world-cup-2019")


from  bs4 import BeautifulSoup

_parser = 'html.parser'
_soup = BeautifulSoup(_url, _parser)
print(_soup.table)