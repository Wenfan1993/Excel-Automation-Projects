# -*- coding: utf-8 -*-
"""
Created on Tue Jul  7 00:01:14 2020

@author: wenfan
"""


import pandas as pd
import os
import sys
import numpy as np
 
def get_clean_dataframe(df:'the raw df',
                        locator:'string, the column name of the 1st column of the df',
                        locator_column=-1)->'clean df':
    if locator_column<0:
        for _, row in df.iterrows():
            for column_num, cell in enumerate(row.values):
                if isinstance(cell, str) and cell.lower().find(locator.lower())>-1:
                    locator_column = column_num
                    break
            if locator_column>=0:
                break   
        else:
            raise ValueError(f"The locator '{locator}' is NOT FOUND in the dataframe. Check locator")

    start_index  = df.loc[df.iloc[:,locator_column].apply(lambda x: x.lower() if isinstance(x, str) else x)==locator.lower()].index[0]
    df_updated = df.iloc[start_index+1:,:]
    df_updated.columns = df.iloc[start_index,:].values
    df_updated = df_updated.reset_index(drop=True)
    
    return df_updated, start_index

def preprocess_df(df:'dataframe, the raw df imported',
                  locator_column=-1,
                  columns_to_keep = None)->'dataframe with headers cleaned and columns not needed removed':

    columns = [item.upper() if isinstance(item, str) else item for item in df.columns]

    try:
        try:        
            df_clean, start_index = get_clean_dataframe(df,
                                          'date',
                                          locator_column)
        except ValueError:
            if 'DATE' in columns:
                df_clean = df.copy()
            else:
                raise ValueError(f'DATE not in the columns of the df')

        df_clean.columns = [header.upper() if isinstance(header, str) else header for header in df_clean.columns]
        df_clean['YEAR'] = df_clean['DATE'].apply(lambda x: x.year)
        df_clean['MONTH'] = df_clean['DATE'].apply(lambda x: x.month)
            
    except ValueError:

        try:
            df_clean, start_index = get_clean_dataframe(df,
                                          'year',
                                          locator_column)
        except ValueError:
            if 'YEAR' in columns:
                df_clean = df.copy()
                pass

        df_clean.columns = [header.upper() if isinstance(header, str) else header for header in df_clean.columns]

    df_clean = df_clean.loc[df_clean['YEAR']>2018]                    

    try:
        header_payment = [name for name in df_clean.columns if isinstance (name, str) and name.lower().find('payment')>-1 and name.lower().find('cum')==-1][0]
    except IndexError:
        header_payment = [name for name in df_clean.columns if isinstance (name, str) and name.lower().find('spend')>-1 and name.lower().find('cum')==-1][0]
        
    header_budget = [name for name in df_clean.columns if isinstance (name, str) and name.lower().find('budget')>-1 and name.lower().find('cum')==-1][0]

    df_updated = df_clean[['YEAR','MONTH',header_payment,header_budget]]
    df_updated.columns = ['YEAR','MONTH','PAYMENT','BUDGET']
    
    if columns_to_keep:
        additional_column = [item.upper() for item in columns_to_keep if item.upper() not in df_updated.columns]
        for col in additional_column:
            try:
                df_updated = pd.concat([df_updated,df_clean[col]], axis=1)
            except KeyError:
                if isinstance(col, str) and col.upper.find('ROW_NUMBER')>-1:
                    df_updated = pd.concat([df_updated,df_clean.iloc[:,-1]], axis=1)
                df_updated.columns = list(df_updated.columns)[:-1]+['ROW_NUMBER']
    
    return df_updated
    
def get_desired_df(df:'dataframe, with year, month, payment and budget',
                   dept:'string, the department id')->'dataframe in line with report':
    df_sum_payment = df.groupby(['YEAR','MONTH']).sum().loc[:,['PAYMENT']]
    df_max_budget = df.groupby(['YEAR','MONTH']).max().loc[:,['BUDGET']]
    df_total = df_sum_payment.join(df_max_budget)
    df_total['DIFFERENCE'] = df_total['BUDGET'] - df_total['PAYMENT']
    df_total.reset_index(inplace = True, drop=False)
    df_total['CUMULATIVE PAYMENT'] =df_total.groupby(['YEAR'])['PAYMENT'].cumsum()
    df_total['CUMULATIVE BUDGET'] =df_total.groupby(['YEAR'])['BUDGET'].cumsum()
    df_total['CUMULATIVE DIFFERENCE'] = df_total['CUMULATIVE BUDGET'] - df_total['CUMULATIVE PAYMENT']
    df_total['PAYMENT ROLLING 3 PERIODS'] = df_total['PAYMENT'].rolling(3).sum()
    df_total['BUDGET ROLLING 3 PERIODS'] = df_total['BUDGET'].rolling(3).sum()
    df_total.fillna(0, inplace = True)
    df_total.insert(0, column='DEPARTMENT ID', value= dept)

    return df_total

def get_dept_from_filename(filename:'str, the filename')->'str, dept id':
    dept_num=''
    for letter in filename:
        try:
            dept_num = dept_num+ str(int(letter))
        except ValueError:
            continue
    return dept_num
