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

        tempDT.iloc[:, 0] = pd.to_datetime(tempDT.iloc[:, 0], infer_datetime_format=True,errors='coerce')
        tempDT.set_index([tempDT.columns[0]], inplace=True)

        for row_index, row in tqdm(tempDT.iterrows(), desc="数据清洗进度"):
            if not is_numeric_dtype(pd.Index(row)):
                print('移除存在异常数据行:', row_index)
                row.drop(row.index, inplace=True)
                tempDT.loc[row_index] = row
        tempDT.replace(-999, np.nan,inplace=True)
        # print(tempDT)
        return tempDT
