# -*- coding: utf-8 -*-
"""
Created on Tue Nov  8 22:28:15 2022

@author: XIAOQING CHEN
"""

import pandas as pd
#import numpy as np
# import xlwings as xw
# import os
import re
import time
# import datetime
# import shutil
# from openpyxl import load_workbook
# from openpyxl.utils import get_column_letter
# from openpyxl.styles import Font
# from openpyxl.styles import colors
# from openpyxl.styles import Alignment
# from openpyxl.styles import PatternFill
# from openpyxl.styles import Side,Border

#%%
class DataFrameOpe():
    def __init__(self):
        self.system_date = time.strftime("%Y%m%d",time.localtime(time.time())) #系统时间YYYYMMMDD(字符串)

    def dataframe_attributes(self, df:pd.DataFrame):
        self.row_number = df.shape[0]
        self.column_number = df.shape[1]
        
    def reset_index_by_header(self, df:pd.DataFrame):
        """       
        Purpose
        ------------------------
        Reset Header of Dataframe. Set the first row that is not empty as the header.
        ------------------------
        
        Parameters
        ------------------------
        df: Dataframe need to be updated.
        ------------------------
        """
        #find_df_title_index 截取前十行及以内的数据
        find_df_title_index = pd.DataFrame()
        if df.shape[0] >= 10:
            find_df_title_index = df[:10]
        else:
            find_df_title_index = df
        
        #df_title_index 确定第一个不为空的行数
        df_title_index = None
        
        for i in range(find_df_title_index.shape[0]):
            df_value = find_df_title_index[find_df_title_index.columns[1]][i]
            if pd.notnull(df_value):
                df_title_index = i
                break
        
        #df_new 重新定义header以及表格内容
        df_new = df.copy(deep=True)
        
        if df_title_index != None:
            df_col_name = df_new.loc[df_title_index].to_dict()
            df_new = df_new.loc[df_title_index + 1:].rename(columns = df_col_name)
        else:
            print('Please note: Header not found. Output Raw data.\n')
        
        return df_new

    def reset_index_by_patient(self, df:pd.DataFrame):
        """       
        Purpose
        ------------------------
        Reset Header of Dataframe. Set the first row that is not empty as the header.
        ------------------------
        
        Parameters
        ------------------------
        df: Dataframe need to be updated.
        ------------------------
        """
        
        #find_df_title_index 截取前十行及以内的数据
        find_df_title_index = pd.DataFrame()
        if df.shape[0] >= 10:
            find_df_title_index = df[:10]
        else:
            find_df_title_index = df

        #df_title_index, subject_id_col 确认patient所在的行列数
        find_subject_id = False
        df_title_index = 0
        subject_id_col = 0
        
        for i in range(find_df_title_index.shape[1]):
            for j in range(find_df_title_index.shape[0]):
                df_value = str(find_df_title_index.iloc[j,i]).lower().replace(' ','')
                
                if re.search('subject',df_value) and 'status' not in df_value and len(df_value) <= 15:
                    df_title_index = j
                    subject_id_col = i
                    find_subject_id = True
                    break
                elif re.search('patient',df_value) and len(df_value) <= 38:
                    df_title_index = j
                    subject_id_col = i
                    find_subject_id = True
                    break
                elif re.search('ssid',df_value) and len(df_value) <= 10:
                    df_title_index = j
                    subject_id_col = i
                    find_subject_id = True
                    break
                elif df_value == 'pt':
                    df_title_index = j
                    subject_id_col = i
                    find_subject_id = True
                    break

            if find_subject_id == True:
                break

        #df_new 重新定义header以及表格内容
        df_new = df.copy(deep=True)
        col_name = "Patient ID"
        
        if find_subject_id == True:
            df_new.loc[df_title_index,subject_id_col] = col_name
            df_col_name = df_new.loc[df_title_index].to_dict()
            df_new = df_new.loc[df_title_index + 1:].rename(columns = df_col_name)

            df_new[col_name] = df_new[col_name].astype("str")
            df_new[col_name] = df_new[col_name].str.strip() #单列删除空格
        else:
            print('Column contains <Patient> not found. Output Raw data. \n')

        return df_new

