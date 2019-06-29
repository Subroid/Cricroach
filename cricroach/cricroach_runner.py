
from cricroach.stats.stats import StatsFetcher as SFetcher

URL_ = 'https://www.espncricinfo.com/series/8039/game/1144519/australia-vs-new-zealand-37th-match-icc-cricket-world-cup-2019'
FOLDER_LOC_ = 'E:/D11/'
sf_ = SFetcher()
sf_.full_fetch(URL_, FOLDER_LOC_)