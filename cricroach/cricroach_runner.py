
from cricroach.stats.stats import StatsFetcher as SFetcher

URL_ = 'https://www.espncricinfo.com/series/8039/game/1144530/england-vs-new-zealand-final-icc-cricket-world-cup-2019'
FOLDER_LOC_ = 'E:/D11/'
sf_ = SFetcher()
sf_.full_fetch(URL_, FOLDER_LOC_)