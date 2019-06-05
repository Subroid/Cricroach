
from cricfetcher.stats.stats import StatsFetcher as SFetcher

URL_ = 'http://stats.espncricinfo.com/ci/engine/player/253802.html?class=2;template=results;type=batting'
FOLDER_LOC_ = 'E:/D11/D11/'
sf_ = SFetcher()
sf_.full_fetch(URL_, FOLDER_LOC_)