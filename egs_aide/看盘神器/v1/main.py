# -*- coding: utf-8 -*-
# @Author   : liyi
# @Time     : 2023/1/21 23:05
# @File     : edit_active_excel.py
# @Project  : main
# Copyright (c) Personal 2022 liyi
# Function Description: 
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import os
import random
import time
import traceback
from typing import Tuple

import xlwings as xw
import pandas as pd
import qstock as qs

import logging
# import easytrader


class StockMonitor:
    @staticmethod
    def normalize_stock_code(code) -> str:
        if pd.isna(code):
            return ''

        if isinstance(code, (int, float)):
            if isinstance(code, float) and code.is_integer():
                code_str = str(int(code))
            else:
                code_str = str(code)
        else:
            code_str = str(code).strip()

        if not code_str or code_str.lower() in ['nan', 'none']:
            return ''

        if '.' in code_str:
            parts = code_str.split('.')
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isalpha():
                code_str = parts[0]

        if code_str.isdigit() and len(code_str) <= 6:
            code_str = code_str.zfill(6)

        return code_str

    def __init__(self,
                 monitor_xlsx_file: str,
                 stock_xlsx_file: str,
                 data_api: str,
                 refresh_wait_time: int,
                 ):
        # 1. open monitor xlsx file
        if not os.path.exists(monitor_xlsx_file):
            logging.info('文件路径错误或不存在：' + monitor_xlsx_file)

        self.wb = xw.Book(monitor_xlsx_file)
        self.sht_num = len(self.wb.sheets)

        # 2. open stock info xlsx file
        self.stock_xlsx_file = stock_xlsx_file
        self.use_online_data = False   # if True, use qstock returned dataframe
        self.static_properties_lst = ['证券代码', '证券简称', '所属热门概念', '所属概念板块', '所属Wind行业名称',
                                      '所属申万行业名称(2021)', '所属产业链板块',
                                      '所属行政区划[行政区划级别]省级', '公司属性', '所属规模风格类型']
        self.df_stock = pd.DataFrame()
        self.load_stock_info_from_excel()

        # ===== Internal Variables =======
        self._wb_dict = {}
        self._refresh_wait_time = refresh_wait_time
        self._api_min_interval = {
            'watchlist': 1,
            'concept': 1,
            'billboard': 1,
            'news': 1,
            'zt': 1,
        }
        self._api_next_allowed = {k: 0.0 for k in self._api_min_interval}
        self._api_fail_count = {k: 0 for k in self._api_min_interval}

        # user = easytrader.use('universal_client') # universal_client 支持多个券商，需安装对应券商的客户端并登录


    @staticmethod
    def update_sheet_data_only(sheet, df_sht: pd.DataFrame, row_num: int, col_num: int):
        """
        仅更新数据区(不重写表头)，尽量保持原有 Excel 格式不变。
        约定第 1 行为表头，数据从第 2 行开始。
        """
        if df_sht is None or df_sht.empty:
            return

        data_rows_in_sheet = max(row_num - 1, 0)
        if data_rows_in_sheet <= 0 or col_num <= 0:
            return

        write_row_num = min(data_rows_in_sheet, len(df_sht))
        write_col_num = min(col_num, len(df_sht.columns))
        if write_row_num <= 0 or write_col_num <= 0:
            return

        data_values = df_sht.iloc[:write_row_num, :write_col_num].values.tolist()
        sheet.range((2, 1), (write_row_num + 1, write_col_num)).value = data_values

    @staticmethod
    def update_df_existing_columns(target_df: pd.DataFrame,
                                   source_df: pd.DataFrame,
                                   col_map: dict = None) -> bool:
        """
        仅更新 target_df 已有列，不新增任何列。
        :param target_df: 目标表（来自现有 sheet）
        :param source_df: 数据源表（接口返回）
        :param col_map: 可选列映射，格式 {target_col: source_col}
        :return: 是否有列被更新
        """
        if target_df is None or source_df is None or target_df.empty or source_df.empty:
            return False

        target_len = len(target_df)
        updated = False

        if col_map is None:
            pairs = [(col, col) for col in target_df.columns if col in source_df.columns]
        else:
            pairs = [(t_col, s_col)
                     for t_col, s_col in col_map.items()
                     if t_col in target_df.columns and s_col in source_df.columns]

        for target_col, source_col in pairs:
            src_series = source_df[source_col].reset_index(drop=True)
            aligned_series = src_series.reindex(range(target_len))
            target_df[target_col] = aligned_series.values.tolist()
            updated = True

        return updated

    @staticmethod
    def sort_news_by_datetime_desc(df_news: pd.DataFrame) -> pd.DataFrame:
        """
        按新闻日期+时间降序排列（最新在前）。
        兼容列名：发布日期 + 发布时间/时间。
        """
        if df_news is None or df_news.empty:
            return df_news

        date_col = '发布日期' if '发布日期' in df_news.columns else None
        time_col = None
        if '发布时间' in df_news.columns:
            time_col = '发布时间'
        elif '时间' in df_news.columns:
            time_col = '时间'

        if date_col is None and time_col is None:
            return df_news

        df_sorted = df_news.copy()

        if date_col and time_col:
            dt_series = pd.to_datetime(
                df_sorted[date_col].astype(str).str.strip() + ' ' +
                df_sorted[time_col].astype(str).str.strip(),
                errors='coerce'
            )
            df_sorted['_sort_dt'] = dt_series
            df_sorted.sort_values(by='_sort_dt', ascending=False, inplace=True)
            df_sorted.drop(columns=['_sort_dt'], inplace=True)
        elif date_col:
            date_series = pd.to_datetime(df_sorted[date_col], errors='coerce')
            df_sorted['_sort_dt'] = date_series
            df_sorted.sort_values(by='_sort_dt', ascending=False, inplace=True)
            df_sorted.drop(columns=['_sort_dt'], inplace=True)
        else:
            time_series = pd.to_datetime(df_sorted[time_col].astype(str), errors='coerce')
            df_sorted['_sort_dt'] = time_series
            df_sorted.sort_values(by='_sort_dt', ascending=False, inplace=True)
            df_sorted.drop(columns=['_sort_dt'], inplace=True)

        return df_sorted.reset_index(drop=True)

    def can_request_api(self, api_name: str) -> bool:
        return time.time() >= self._api_next_allowed.get(api_name, 0.0)

    def load_stock_info_from_excel(self):
        if not os.path.exists(self.stock_xlsx_file):
            self.df_stock = pd.DataFrame()
            self.use_online_data = True
            return

        try:
            df_stock = pd.read_excel(self.stock_xlsx_file, header=0, index_col=None)
            for item in self.static_properties_lst:
                for col_name in df_stock.columns:
                    col_name_tmp = col_name.replace('\n', '').replace(' ', '')
                    if item in col_name_tmp:
                        df_stock.rename(columns={col_name: item}, inplace=True)
                        break

            self.df_stock = df_stock[self.static_properties_lst].copy()
            self.df_stock['证券代码'] = self.df_stock['证券代码'].apply(self.normalize_stock_code)
            self.use_online_data = False
        except Exception as e:
            logging.error('读取个股信息Excel失败: %s', e)
            self.df_stock = pd.DataFrame()
            self.use_online_data = True

    def mark_api_result(self, api_name: str, success: bool):
        now = time.time()
        min_interval = self._api_min_interval.get(api_name, 5)

        if success:
            self._api_fail_count[api_name] = 0
            jitter = random.uniform(0, max(min_interval * 0.2, 0.2))
            self._api_next_allowed[api_name] = now + min_interval + jitter
        else:
            fail_count = self._api_fail_count.get(api_name, 0) + 1
            self._api_fail_count[api_name] = fail_count
            backoff = min(300, min_interval * (2 ** min(fail_count, 6)))
            jitter = random.uniform(0, max(backoff * 0.2, 0.5))
            self._api_next_allowed[api_name] = now + backoff + jitter
            logging.warning('接口%s触发退避，第%d次失败，%.1f秒后重试', api_name, fail_count, backoff + jitter)

    # ========= Funcs ===========
    @staticmethod
    def get_stock_lst(df_sht: pd.DataFrame, remove_postfix: bool = False) -> list:
        stock_raw_lst = df_sht['代码'].tolist()
        stock_lst = []
        for code in stock_raw_lst:
            code_str = StockMonitor.normalize_stock_code(code)
            if not code_str:
                continue

            if remove_postfix:
                code_str = code_str.split('.')[0]

            stock_lst.append(code_str)

        return stock_lst

    def sheet_2_df(self, index: int):
        """
        excel sheet to dataframe
        :param index: index of sheet
        :return:
        """
        # loading sheet and get basic info
        sheet = self.wb.sheets[index]
        # log.info('Processing: ' + sheet.name)

        # NOTE:
        # - On Windows (pywin32), `sheet.api.UsedRange` is available.
        # - On macOS (appscript), the same attribute access may fail.
        # Use xlwings' cross-platform abstraction instead.
        used_range = sheet.used_range
        row_num = used_range.last_cell.row
        col_num = used_range.last_cell.column
        if row_num == 1 and col_num == 1:
            df_sht = pd.DataFrame()
        else:
            # convert sheet to dataframe
            # todo: headers and index, consider as a set option
            df_sht = sheet.range((1, 1), (row_num, col_num)). \
                options(pd.DataFrame, headers=True, index=False).value

        return sheet, df_sht, row_num, col_num

    # ========= Main Process ===========
    def query_static_info(self):
        print('初始化信息加载。。。')
        for i in range(self.sht_num):
            sheet, df_sht, row_num, col_num = self.sheet_2_df(i)
            if df_sht.empty and 'sheet' in sheet.name.lower():
                continue

            if '自选股' in sheet.name and not self.use_online_data:
                stock_lst = self.get_stock_lst(df_sht, remove_postfix=False)
                if not len(stock_lst):
                    continue

                # get stock info from full table
                df_rows = []
                for code in stock_lst:
                    df_row = self.df_stock[self.df_stock['证券代码'] == code]
                    df_rows.append(df_row)

                if df_rows:
                    df_tmp = pd.concat(df_rows, ignore_index=True)
                else:
                    df_tmp = pd.DataFrame(columns=self.df_stock.columns)

                # put info into sheet table
                for col in df_tmp.columns:
                    if col in df_sht.columns:
                        df_sht[col] = df_tmp[col]

            # update to excel online
            self.update_sheet_data_only(sheet, df_sht, row_num, col_num)

            self._wb_dict[i] = {'sheet': sheet, 'df_sht': df_sht,
                                'row_num': row_num, 'col_num': col_num}

    def query_rt_info(self):
        self.load_stock_info_from_excel()

        for i in range(self.sht_num):
            sheet, df_sht, row_num, col_num = self.sheet_2_df(i)
            if df_sht.empty:
                continue

            if '自选股' in sheet.name:
                if not self.can_request_api('watchlist'):
                    continue

                print('加载自选股数据：', sheet.name)
                # get stock real time data
                stock_lst = self.get_stock_lst(df_sht, remove_postfix=True)
                if not len(stock_lst):
                    continue

                # 1. 获取实时数据，数据来源东方财富
                try:
                    df_rt = qs.realtime_data(code=stock_lst)  # 获取沪深A股最新行情指标
                    self.mark_api_result('watchlist', success=True)
                except Exception as e:
                    logging.error('Caught exception in realtime Data Acquisition %s' % e)
                    traceback.print_exc()
                    self.mark_api_result('watchlist', success=False)
                    df_rt = pd.DataFrame()

                # 2. 获取交易日实时盘口异动数据，相当于盯盘小精灵
                # df_chg = qs.realtime_change()
                # print('实时盘口异动数据：')
                # print(df_chg)

                # 使用新闻统一接口，无数据会报错
                # df_stock_news = qs.stock_news('天瑞仪器')

                # store data in dataframe
                if not df_rt.empty:
                    if not self.use_online_data:
                        rt_col_map = {
                            '现价(元)': '最新',
                            '涨跌幅': '涨幅',
                            '刷新时间': '时间',
                        }
                        self.update_df_existing_columns(df_sht, df_rt, rt_col_map)
                        # inplace: 原地修改
                        # df_sht.sort_values(by="涨跌幅", inplace=True, ascending=False)
                    else:
                        self.update_df_existing_columns(df_sht, df_rt)

                    # update to excel online
                    self.update_sheet_data_only(sheet, df_sht, row_num, col_num)

            if '概念涨幅榜' in sheet.name:
                if not self.can_request_api('concept'):
                    continue

                print('加载概念涨幅榜数据：', sheet.name)
                try:
                    df_concept = qs.realtime_data('概念板块')  # 获取概念板块最新行情指标: 来源东方财富
                    self.mark_api_result('concept', success=True)
                    if not df_concept.empty:
                        # sheet.range((1, 1), df_concept.shape).value = df_concept
                        if self.update_df_existing_columns(df_sht, df_concept):
                            self.update_sheet_data_only(sheet, df_sht, row_num, col_num)
                except Exception as e:
                    self.mark_api_result('concept', success=False)
                    logging.error('Caught exception in realtime concept Data Acquisition %s' % e)
                    continue

            if '龙虎榜' in sheet.name:
                if not self.can_request_api('billboard'):
                    continue

                print('加载龙虎榜数据：', sheet.name)
                try:
                    df_head = qs.stock_billboard()  # 获取龙虎榜最新行情指标: 来源东方财富
                    self.mark_api_result('billboard', success=True)
                    if not df_head.empty:
                        if self.update_df_existing_columns(df_sht, df_head):
                            self.update_sheet_data_only(sheet, df_sht, row_num, col_num)
                except Exception as e:
                    self.mark_api_result('billboard', success=False)
                    logging.error('Caught exception in billboard Data Acquisition %s' % e)
                    traceback.print_exc()
                    continue

            if '财联社新闻' in sheet.name:
                if not self.can_request_api('news'):
                    continue

                print('加载财联社新闻：', sheet.name)
                try:
                    df_news = qs.news_data()  # 获取财联社新闻
                    self.mark_api_result('news', success=True)
                    if not df_news.empty:
                        df_news = self.sort_news_by_datetime_desc(df_news)

                        if '发布时间' in df_news.columns:
                            df_news['发布时间'] = df_news['发布时间'].apply(str)
                        if '发布日期' in df_news.columns:
                            df_news['发布日期'] = df_news['发布日期'].apply(str)

                        if self.update_df_existing_columns(df_sht, df_news):
                            self.update_sheet_data_only(sheet, df_sht, row_num, col_num)
                except Exception as e:
                    self.mark_api_result('news', success=False)
                    logging.error('Caught exception in Finance News Acquisition %s' % e)
                    traceback.print_exc()
                    continue

            # if '市场快讯' in sheet.name:
            #     print('加载市场快讯：', sheet.name)
            #     try:
            #         df_js = qs.news_data('js')  # 获取市场快讯
            #         if not df_js.empty:
            #             sheet.range((1, 1), df_js.shape).value = df_js
            #     except Exception as e:
            #         logging.error('Caught exception in Market Express Acquisition %s' % e)
            #         traceback.print_exc()
            #         continue

            if '涨停板' in sheet.name:
                if not self.can_request_api('zt'):
                    continue

                print('加载涨停板：', sheet.name)
                try:
                    df_zt = qs.stock_zt_pool()
                    self.mark_api_result('zt', success=True)
                    if not df_zt.empty:
                        if self.update_df_existing_columns(df_sht, df_zt):
                            self.update_sheet_data_only(sheet, df_sht, row_num, col_num)
                except Exception as e:
                    self.mark_api_result('zt', success=False)
                    logging.error('Caught exception in Market Express Acquisition %s' % e)
                    traceback.print_exc()
                    # continue

            pass

    def real_time_update(self):
        self.query_static_info()
        while True:
            self.query_rt_info()
            time.sleep(self._refresh_wait_time)


def main():
    # 1. params
    monitor_xlsx_file = os.path.join(os.getcwd(), '看盘模板.xlsx')  # 看盘模板
    stock_xlsx_file = os.path.join(os.getcwd(), '全部A股信息.xlsx')
    data_api = 'qstock'  # wind, qstock
    refresh_wait_time = 1  # 刷新等待时间(秒)

    # 2. process
    monitor = StockMonitor(monitor_xlsx_file, stock_xlsx_file, data_api, refresh_wait_time)
    monitor.real_time_update()


if __name__ == '__main__':
    main()
