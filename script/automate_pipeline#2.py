# -*- coding: utf-8 -*-

import pandas as pd
import os
from datetime import datetime 
from openpyxl.utils import get_column_letter
import win32com.client
import sys

path_script = r'C:\Users\Wenxi\Desktop\project#2\Class#0\script'
if path_script not in sys.path:
    sys.path.append(path_script)

root = r'C:\Users\Wenxi\Desktop\project#2\Class#0'
os.chdir(root)


from func_absolute_value import *
from func_formula import *

#define if you would like to use formula or not
use_formula = True

os.chdir('.\input') 
files = os.listdir()

sub_file_list = []
#create a dictionary for storing the file : department number  
dept_num_dict = {}

for file in files:
    if file.lower().find('dep')>-1 and file.find('$')==-1:
        #get the department number
        dept_num = get_dept_from_filename(file)
        dept_num_dict[file] = dept_num
        
        #locate the dataframe 
        dept_df_orig = pd.read_excel(file)
        
        if use_formula:
            #pre-process the dataframe to get the payment and budget            
            output_df = get_processed_df_with_formula(dept_df_orig,
                                                      dept_num)
            sub_file_list.append(output_df)
            
        else:
            #pre-process the dataframe to get the payment and budget
            dept_df = preprocess_df(dept_df_orig,
                                    -1)
            #process to get the update the dataframe in the required format
            dept_df_updated = get_desired_df(dept_df,
                                             dept_num)
            
            #collect the processed dataframes by appending it to the list object        
            sub_file_list.append(dept_df_updated)

report_total = pd.concat(sub_file_list)
report_output = report_total.sort_values(by = [*report_total.columns[:3]])


os.chdir(os.path.join(root, 'output'))
if use_formula:
    #add the cumulative columns and rolling columns                
    report_output = add_cum_rolling_columns(report_output)

    with pd.ExcelWriter(f'Report_{datetime.now().month}_{datetime.now().day}.xlsx') as writer:
        report_output.to_excel(writer, sheet_name = 'Summary', index=False)
        for file in files:
            if file.lower().find('dep')>-1 and file.find('$')==-1:
                pd.read_excel(os.path.join(root,'input',file)).to_excel(writer, sheet_name = f'Support Dept {dept_num_dict[file]}', index=False)
    
else:    
    report_output.to_excel(f'Report_{datetime.now().month}_{datetime.now().day}.xlsx', index=False)

    
