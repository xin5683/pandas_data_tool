import datetime
import time
import os
from glob import iglob
from natsort import natsorted
import numpy as np
import pandas as pd
from pandas.api.types import is_numeric_dtype
import openpyxl
import xlrd
from tqdm import tqdm
from dateutil.parser import parse


def validate_datetime(date_text):
    try:
        parse(date_text)
        return True
    except ValueError:
        return False


class DataTools:

    def __init__(self, excelPath):
        self.sourcePd = pd.DataFrame(self.__clean_data(excelPath))

    @staticmethod
    def __clean_data(excelPath):
        """
        :param excelPath:
        :return:
        """
        tempDT = pd.read_excel(excelPath, sheet_name=0)
        # 识别第一行数据是否是其他数据清除
        if not validate_datetime(str(tempDT.loc[0][0])):
            tempDT.drop(tempDT.index[0], inplace=True)

        tempDT.iloc[:, 0] = pd.to_datetime(tempDT.iloc[:, 0], infer_datetime_format=True, errors='coerce')
        tempDT.set_index([tempDT.columns[0]], inplace=True)

        for row_index, row in tqdm(tempDT.iterrows(), desc="Data cleaning progress"):
            if not is_numeric_dtype(pd.Index(row)):
                print('Remove abnormal data:', row_index)
                row.drop(row.index, inplace=True)
                tempDT.loc[row_index] = row
        tempDT.replace(-999, np.nan, inplace=True)
        # print(tempDT)
        return tempDT

    @staticmethod
    def __check_data(Series, check_list):
        count_num = 0
        user_Series = pd.Series(Series, dtype='float64')
        for element in check_list:
            if element in user_Series.index:
                if user_Series[element] == -999:
                    count_num += 1
                    if count_num > 4:
                        break
        if count_num > 4:
            return False
        else:
            return True

    def __effective_rate_clean_data(self):
        elements = ['乙烷', '乙烯', '丙烷', '丙烯', '异丁烷', '乙炔', '正丁烷', '异戊烷', '正戊烷', '氟利昂11', '异戊二烯', '氟利昂113', '丙酮', '四氯化碳',
                    '苯',
                    '甲苯', '乙基苯', '间/对二甲苯', '邻二甲苯', '苯乙烯']
        dataFrame = pd.DataFrame(self.sourcePd)
        for row_index, row in dataFrame.iterrows():
            if not self.__check_data(row, elements):
                dataFrame.loc[row_index] = pd.Series([-999 for x in range(dataFrame.shape[1])], index=row.index.values,
                                                     dtype='float64')
        return dataFrame

    def get_total_effective_rate(self):
        """
        总有效率
        :return:总有效率
        """
        column_list = []
        dataFrame = pd.DataFrame(self.__effective_rate_clean_data()).T
        user_list = pd.Series(dtype='float64')
        for row_index, row in tqdm(dataFrame.iterrows(), desc='总有效率计算'):
            value_counts = row.value_counts(dropna=False)
            # print(value_counts.index[0], type(value_counts.index[0]))
            if np.nan in value_counts.index:
                user_list[row_index] = (row.size - value_counts[np.nan]) / row.size
            else:
                user_list[row_index] = 1.0
        column_list.append({'column': "总有效率", 'data': user_list})
        return column_list

    def get_week_effective_rate(self):
        week_freq = ('W-SUN', 'W-MON', 'W-TUE', 'W-WED', 'W-THU', 'W-THU', 'W-FRI', 'W-SAT')
        weeks = [g for n, g in
                 self.__effective_rate_clean_data().groupby(
                     pd.Grouper(freq=week_freq[pd.Series(self.__effective_rate_clean_data().index[0]).dt.weekday[0]]))]
        user_list = pd.Series(dtype='float64')
        column_list = []
        for week in tqdm(weeks, desc="周有效率计算"):
            col_name = week.index[0].strftime('%Y/%m/%d') + '-' + week.index[len(week.index) - 1].strftime('%Y/%m/%d')
            dataFrame = pd.DataFrame(week).T
            for row_index, row in dataFrame.iterrows():
                value_counts = row.value_counts(dropna=False)
                if np.nan in value_counts.index:
                    user_list[row_index] = (row.size - value_counts[np.nan]) / row.size
                else:
                    user_list[row_index] = 1.0
            column_list.append({'column': col_name, 'data': user_list})
        return column_list

    def get_effective_rate(self):
        """
        获取有效率
        :return: 返回 DataFrame
        """
        effective_rate_list = self.get_total_effective_rate() + self.get_week_effective_rate()
        dataFrame = pd.DataFrame()
        for data in effective_rate_list:
            dataFrame[data['column']] = data['data']
        # print(dataFrame)
        return dataFrame
