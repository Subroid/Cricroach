
from cricprep.top_performances import TopPerformancesFetcher as TPFetcher

URL_ = 'https://www.espncricinfo.com/series/19408/game/1193504/sri-lanka-vs-bangladesh-1st-odi-bangladesh-in-sri-lanka-2019'
FOLDER_LOC_ = 'E:/D11/'

tpf_ = TPFetcher()
tpf_.get_head_to_head_last_5_matches_links(URL_)