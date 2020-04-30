#   Copyright (c)  2020. Hstar
#
#  Permission is hereby granted, free of charge, to any person
#  obtaining a copy of this software and associated documentation
# #  files (the "Software"), to deal in the Software without restriction,
# #  including without limitation the rights to use, copy, modify, merge,
# #  publish, distribute, sublicense, and/or sell copies of the Software,
# #  and to permit persons to whom the Software is furnished to do so,
# #  subject to the following conditions:
#
#  The above copyright notice and this permission notice shall be
#  included in all copies or substantial portions of the Software.
#
#  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
#  EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
#  OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE
#  AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
#  HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
#  WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
#  FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
#  OTHER DEALINGS IN THE SOFTWARE.

import datetime
import time
import os
import numpy as np
import pandas as pd
from pandas.api.types import is_numeric_dtype
import openpyxl
import xlrd
from tqdm import tqdm
from dateutil.parser import parse
import os.path
import tkinter as tk
from tkinter import filedialog


def validate_datetime(date_text):
    try:
        parse(date_text)
        return True
    except ValueError:
        return False


def get_now_time():
    import time
    ft_time = time.strftime('%Y%m%d_%H%M', time.localtime(time.time()))
    return ft_time


elements = ['乙烷', '乙烯', '丙烷', '丙烯', '异丁烷', '乙炔', '正丁烷', '异戊烷', '正戊烷', '氟利昂11',
            '异戊二烯', '氟利昂113', '丙酮', '四氯化碳', '苯', '甲苯', '乙基苯', '间/对二甲苯', '邻二甲苯', '苯乙烯']


class DataTools:

    def __init__(self, excelPath=None):
        if excelPath:
            self.source_excelPath = excelPath;
        else:
            root = tk.Tk()
            root.withdraw()
            self.source_excelPath = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('EXCEL', '*.xlsx')])
            if self.source_excelPath is '':
                print("文件打开失败，请选择文件！")
                exit(-2)
        self.material_attribute_path = './116物质属性.xlsx'
        self.attribute_dt = pd.read_excel(self.material_attribute_path, sheet_name=0, keep_default_na=True,
                                          na_values='')
        self.attribute_dt.replace('', np.nan, inplace=True)
        # print(self.attribute_dt[self.attribute_dt['目标化合物名称'] == '异戊烷'])
        self.sourcePd = pd.DataFrame(self.__clean_data(self.source_excelPath, self.attribute_dt))

    def __get_base_data(self, Index, Column):
        return self.attribute_dt[self.attribute_dt['目标化合物名称'] == Index][Column]

    # def __get_material_attribute(self):
    #     attribute_dt = pd.read_excel(self.material_attribute_path, sheet_name=0)
    #     # print(list(attribute_dt[attribute_dt['分类'] == 1]['目标化合物名称']))
    #     # print(pd.Series(attribute_dt.loc[:,'分类'].value_counts()))
    #     # print(attribute_dt.loc[attribute_dt.loc[:, '分类' == 1]])
    def get_unit_conversion(self):
        """
        单位换算
        :return:DataFrame
        """
        dataFrame = pd.DataFrame(self.sourcePd)
        tmp_dataFrame = pd.DataFrame()
        for row_index, row in tqdm(dataFrame.iteritems(), desc="单位换算"):
            data_series = row * self.attribute_dt[self.attribute_dt['目标化合物名称'] == row_index]['分子量'].values[0] / 22.4
            tmp_dataFrame[row_index] = data_series
        # print(tmp_dataFrame)
        return tmp_dataFrame

    def get_uncertainty(self):
        """
        获取不确定度
        确定度 = (((0.1*浓度)) * 浓度) ^ 2 + (0.5 * 检出限) ^ 2) ^ 0.5
        不确定度 = 检出限 * 5 / 6
        :return: DataFrame
        """
        dataFrame = pd.DataFrame(self.get_unit_conversion())
        tmp_dataFrame = pd.DataFrame()
        for row_index, row in tqdm(dataFrame.iteritems(), desc="不确定度"):
            tmp1 = row[row > self.__get_base_data(row_index, '检出限（微克）').values[0]]
            tmp2 = row[row <= self.__get_base_data(row_index, '检出限（微克）').values[0]]
            if len(tmp2):
                tmp2.values[:] = self.__get_base_data(row_index, '检出限（微克）').values[0] * 5 / 6

            s = pd.concat([row[row.isnull()],
                           (((tmp1 * 0.1) * tmp1) ** 2 + (
                                   (0.5 * self.__get_base_data(row_index, '检出限（微克）').values[0]) ** 2)) ** 0.5,
                           tmp2
                           ])
            s.sort_index(inplace=True)
            tmp_dataFrame[s.name] = s
        # print(tmp_dataFrame)
        return tmp_dataFrame

    def get_classify_sum(self, ext_DataFrame=pd.DataFrame()):
        """
        浓度分类和
        :return: DataFrame
        """
        classify_s = pd.Series(self.attribute_dt.loc[:, '分类'].value_counts())
        # print(ext_DataFrame.empty)
        # exit(0)
        if ext_DataFrame.empty:
            dataFrame = pd.DataFrame(self.sourcePd)
        else:
            dataFrame = pd.DataFrame(ext_DataFrame)

        tmp_dataFrame = pd.DataFrame()
        classify_s.sort_index(inplace=True)
        # print(classify_s)
        for items in classify_s.iteritems():
            data_series = dataFrame[list(self.attribute_dt[self.attribute_dt['分类'] == items[0]]['目标化合物名称'])].sum(axis=1)
            tmp_dataFrame[str(items[0])] = data_series
        tmp_dataFrame['T'] = tmp_dataFrame.sum(axis=1)
        # print(tmp_dataFrame)
        return tmp_dataFrame

    def get_OFP(self):
        dataFrame = pd.DataFrame(self.sourcePd)
        tmp_dataFrame = pd.DataFrame()
        # print(dataFrame)
        for row_index, row in tqdm(dataFrame.iteritems(), desc="OFP progress"):
            tmp_dataFrame[row_index] = row * self.attribute_dt[self.attribute_dt['目标化合物名称'] == row_index]['分子量'].values[
                0] * self.attribute_dt[self.attribute_dt['目标化合物名称'] == row_index]['MIR'].values[0] / 22.4
        # print(tmp_dataFrame)
        return tmp_dataFrame

    def get_OFP_classify_sum(self):
        """
        OFP的分类和
        :return:
        """
        dataFrame = pd.DataFrame(self.get_OFP())
        # print(dataFrame)
        return self.get_classify_sum(ext_DataFrame=dataFrame)

    @staticmethod
    def __clean_data(excelPath, attribute_dt):
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
        # 删除不在统计范围内的列

        drop_list = list(set(list(tempDT.columns)) ^ set(list(attribute_dt['目标化合物名称'])))
        # print(tempDT)
        tempDT.drop(drop_list, inplace=True, axis=1)

        # print(tempDT)
        return tempDT

    @staticmethod
    def __check_data(Series, check_list):
        count_num = 0
        user_Series = pd.Series(Series, dtype='float64')
        for element in check_list:
            if element in user_Series.index:
                if user_Series[element] == np.nan:  # 数据清洗会将-999都变为nan,
                    count_num += 1
                    if count_num > 4:
                        break
        if count_num > 4:
            return False
        else:
            return True

    def __effective_rate_clean_data(self):

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
        """
        周有效率
        :return:
        """
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

        dataFrame.loc['关键物质'] = dataFrame.loc[elements, :].mean(axis=0)
        dataFrame.loc['其他物质'] = dataFrame.loc[list(set(dataFrame.index.values) ^ set(elements)), :].mean(axis=0)

        # print(dataFrame)
        return dataFrame

    def get_SOA(self):
        dataFrame = pd.DataFrame(self.get_unit_conversion())
        SOA_dataFrame = pd.DataFrame()
        # print(dataFrame)
        for row_index, row in tqdm(dataFrame.iteritems(), desc="SOA progress"):
            FAC = self.attribute_dt[self.attribute_dt['目标化合物名称'] == row_index]['FAC']
            FVOC = self.attribute_dt[self.attribute_dt['目标化合物名称'] == row_index]['FVOC']
            if pd.notna(FAC.values[0]) and pd.notna(FVOC.values[0]):
                SOA_dataFrame[row_index] = row * FAC.values[0] / (1 - FVOC.values[0])

        # print(SOA_dataFrame)
        return SOA_dataFrame

    def default_output_all(self):
        if self.source_excelPath is None:
            print("文件路径异常!!!", self.source_excelPath)
            exit(-1)
        out_path = os.path.splitext(self.source_excelPath)[0] + '_' + get_now_time() + \
                   os.path.splitext(self.source_excelPath)[-1]
        start = time.time()
        write = pd.ExcelWriter(out_path)
        self.get_OFP_classify_sum().to_excel(write, sheet_name='OFP分类和', index=True)
        self.get_OFP().to_excel(write, sheet_name='OFP', index=True)
        self.get_classify_sum().to_excel(write, sheet_name='分类和', index=True)
        self.get_unit_conversion().to_excel(write, sheet_name='单位换算', index=True)
        self.get_uncertainty().to_excel(write, sheet_name='不确定度', index=True)
        self.get_effective_rate().to_excel(write, sheet_name='有效率', index=True)
        self.get_SOA().to_excel(write, sheet_name='SOA', index=True)
        print("处理完成正在写入文件...")
        write.save()
        end = time.time()
        print("输出文件：", out_path)
        print('处理时间' + str(datetime.timedelta(seconds=end - start)))
