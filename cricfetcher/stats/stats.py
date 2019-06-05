
class StatsFetcher:
    import urllib.request as Ureq
    import pandas as Pnds
    from pandas import DataFrame as DFrame
    from openpyxl import Workbook as Wb
    from openpyxl.utils import dataframe as odf__
    import re as re__

    def __init__(self):
        self

    def full_fetch(self, match_url, folder_loc):
        df_list_batting_career_summary_stats_ = self.__fetch_career_summaries(match_url)
        df_batting_career_overview_stats_ = df_list_batting_career_summary_stats_[0]
        df_batting_career_summary_stats_ = df_list_batting_career_summary_stats_[1]
        df_batting_career_overview_stats_ = self.__numerize_df_stats_upto_(df_batting_career_overview_stats_, 13)
        df_batting_career_summary_stats_ = self.__numerize_df_stats_upto_(df_batting_career_summary_stats_, 13)
        wb_ = self.Wb()
        excel_file_ = folder_loc + 'test batting career summary.xlsx'
        ws_ = wb_.create_sheet('abc')
        deployable_odf_ = self.odf__.dataframe_to_rows(df_batting_career_overview_stats_, index=False, header=True)
        for df_row in deployable_odf_:
            ws_.append(df_row)
        ws_.append([""])
        deployable_odf_ = self.odf__.dataframe_to_rows(df_batting_career_summary_stats_, index=False, header=True)
        for df_row in deployable_odf_:
            ws_.append(df_row)
        sheet_ = wb_['Sheet']
        wb_.remove(sheet_)
        wb_.save(excel_file_)
        wb_.close()

    def __fetch_career_summaries(self, match_url):
        df_career_overview_stats_ = self.__fetch_career_overview(match_url)
        df_career_summary_stats_ = self.__fetch_career_summary(match_url)
        df_merged_career_overview_stats_ = \
            self.__merge_career_overview_(df_career_overview_stats_, df_career_summary_stats_)
        df_merged_career_overview_stats_ = self.__clean_df_stats_(df_merged_career_overview_stats_)
        df_career_summary_stats_ = self.__clean_df_stats_(df_career_summary_stats_)
        return [df_merged_career_overview_stats_, df_career_summary_stats_]

    def __read_html(self, match_url):
        df_list_ = self.Pnds.read_html(match_url)
        return df_list_

    def __fetch_career_overview(self, match_url):
        df_list_ = self.__read_html(match_url)
        df_stats_ = self.DFrame(df_list_[2])
        df_stats_ = df_stats_.set_index('Unnamed: 0')
        return df_stats_

    def __fetch_career_summary(self, match_url):
        df_list_ = self.__read_html(match_url)
        df_stats_ = self.DFrame(df_list_[3])
        df_stats_ = df_stats_.set_index('Grouping')
        return df_stats_

    def __merge_career_overview_(self, df_stats_1_, df_stats_2_):
        df_bat_1st_stats_ = df_stats_2_.loc[['matches batting first']]
        df_bat_2nd_stats_ = df_stats_2_.loc[['matches fielding first']]
        df_won_batting_1st_stats_ = df_stats_2_.loc[['won batting first']]
        df_won_batting_2nd_stats_ = df_stats_2_.loc[['won fielding first']]
        df_lost_batting_1st_stats_ = df_stats_2_.loc[['lost batting first']]
        df_lost_batting_2nd_stats_ = df_stats_2_.loc[['lost fielding first']]
        df_stats_1_ = df_stats_1_ \
            .append(df_bat_1st_stats_, sort=False) \
            .append(df_bat_2nd_stats_, sort=False) \
            .append(df_won_batting_1st_stats_, sort=False) \
            .append(df_won_batting_2nd_stats_, sort=False) \
            .append(df_lost_batting_1st_stats_, sort=False) \
            .append(df_lost_batting_2nd_stats_, sort=False)
        return df_stats_1_

    def __clean_df_stats_(self, df_stats_):
        df_stats_ = df_stats_.replace(to_replace='-', value='')
        df_stats_ = df_stats_.replace(to_replace='\*', value='', regex=True)
        df_stats_ = df_stats_.replace(to_replace='Profile', value='') #todo remove related column
        df_stats_.drop(labels='Unnamed: 15', axis=1)
        df_stats_ = df_stats_.dropna(axis=1, how='all')
        df_stats_ = df_stats_.fillna("")
        print(df_stats_)
        return df_stats_

    def __numerize_df_stats_upto_(self, df_stats_, upto:int):
        df_stats_.iloc[:, -upto:] = df_stats_.iloc[:, -upto:].apply(self.Pnds.to_numeric)
        return df_stats_

