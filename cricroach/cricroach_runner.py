
from cricroach.stats.stats import StatsFetcher as SFetcher

URL_ = 'http://www.espncricinfo.com/series/8039/game/1144494/england-vs-bangladesh-12th-match-icc-cricket-world-cup-2019'
FOLDER_LOC_ = 'E:/D11/D11/new/'
sf_ = SFetcher()
sf_.full_fetch(URL_, FOLDER_LOC_)