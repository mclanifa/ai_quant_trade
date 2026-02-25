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
import time
import traceback
from typing import Tuple

import xlwings as xw
import pandas as pd
import qstock as qs

import logging


class StockMonitor:
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
        self.use_online_data = False   # if True, use qstock returned dataframe
        if os.path.exists(stock_xlsx_file):
            # todo: consider add options into args
            df_stock = pd.read_excel(stock_xlsx_file, header=0, index_col=None)
            # properties to be show in monitor table and only need to be load once at initialization
            self.static_properties_lst = ['证券代码', '证券简称', '所属热门概念', '所属概念板块', '所属Wind行业名称',
                                          '所属申万行业名称(2021)', '所属产业链板块',
                                          '所属行政区划[行政区划级别]省级', '公司属性', '所属规模风格类型']
            for item in self.static_properties_lst:
                for col_name in df_stock.columns:
                    col_name_tmp = col_name.replace('\n', '').replace(' ', '')
                    if item in col_name_tmp:
                        df_stock.rename(columns={col_name: item}, inplace=True)
                        break

            self.df_stock = df_stock[self.static_properties_lst]
        else:
            self.df_stock = pd.DataFrame()
            self.use_online_data = True   # not using above info

        # ===== Internal Variables =======
        self._wb_dict = {}
        self._refresh_wait_time = refresh_wait_time

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

    # ========= Funcs ===========
    @staticmethod
    def get_stock_lst(df_sht: pd.DataFrame, remove_postfix: bool = False) -> list:
        stock_raw_lst = df_sht['代码'].tolist()
        stock_lst = []
        for code in stock_raw_lst:
            if pd.isna(code):
                continue

            if isinstance(code, (int, float)):
                if isinstance(code, float) and code.is_integer():
                    code_str = str(int(code))
                else:
                    code_str = str(code)
            else:
                code_str = str(code).strip()

            if not code_str or code_str.lower() in ['nan', 'none']:
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
        for val in self._wb_dict.values():
            sheet, df_sht, row_num, col_num = val.values()

            if '自选股' in sheet.name:
                print('加载自选股数据：', sheet.name)
                # get stock real time data
                stock_lst = self.get_stock_lst(df_sht, remove_postfix=True)
                if not len(stock_lst):
                    continue

                # 1. 获取实时数据，数据来源东方财富
                try:
                    df_rt = qs.realtime_data(code=stock_lst)  # 获取沪深A股最新行情指标
                except Exception as e:
                    logging.error('Caught exception in realtime Data Acquisition %s' % e)
                    traceback.print_exc()
                    df_rt = pd.DataFrame()

                # 2. 获取交易日实时盘口异动数据，相当于盯盘小精灵
                # df_chg = qs.realtime_change()

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
                        for target_col, src_col in rt_col_map.items():
                            if target_col in df_sht.columns and src_col in df_rt.columns:
                                df_sht[target_col] = df_rt[src_col].tolist()
                        # inplace: 原地修改
                        # df_sht.sort_values(by="涨跌幅", inplace=True, ascending=False)
                    else:
                        common_cols = [col for col in df_sht.columns if col in df_rt.columns]
                        for col in common_cols:
                            df_sht[col] = df_rt[col].tolist()

                    # update to excel online
                    self.update_sheet_data_only(sheet, df_sht, row_num, col_num)

            # if '概念涨幅榜' in sheet.name:
            #     print('加载概念涨幅榜数据：', sheet.name)
            #     try:
            #         df_concept = qs.realtime_data('概念板块')  # 获取概念板块最新行情指标: 来源东方财富
            #         if not df_concept.empty:
            #             sheet.range((1, 1), df_concept.shape).value = df_concept
            #     except Exception as e:
            #         logging.error('Caught exception in realtime concept Data Acquisition %s' % e)
            #         traceback.print_exc()

            if '龙虎榜' in sheet.name:
                print('加载龙虎榜数据：', sheet.name)
                try:
                    df_head = qs.stock_billboard()  # 获取龙虎榜最新行情指标: 来源东方财富
                    if not df_head.empty:
                        sheet.range((1, 1), df_head.shape).value = df_head
                except Exception as e:
                    logging.error('Caught exception in billboard Data Acquisition %s' % e)
                    traceback.print_exc()
                    continue

            # if '财联社新闻' in sheet.name:
            #     print('加载财联社新闻：', sheet.name)
            #     try:
            #         df_news = qs.news_data()  # 获取财联社新闻
            #         df_news['发布时间'] = df_news['发布时间'].apply(str)
            #         df_news['发布日期'] = df_news['发布日期'].apply(str)
            #         if not df_news.empty:
            #             sheet.range((1, 1), df_news.shape).value = df_news
            #     except Exception as e:
            #         logging.error('Caught exception in Finance News Acquisition %s' % e)
            #         traceback.print_exc()
            #         continue

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
                print('加载涨停板：', sheet.name)
                try:
                    df_zt = qs.stock_zt_pool()
                    if not df_zt.empty:
                        sheet.range((1, 1), df_zt.shape).value = df_zt
                except Exception as e:
                    logging.error('Caught exception in Market Express Acquisition %s' % e)
                    traceback.print_exc()
                    continue

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
