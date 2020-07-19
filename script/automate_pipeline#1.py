# -*- coding: utf-8 -*-
import pandas as pd
import os
from datetime import datetime 
import sys

root = r'C:\Users\Wenxi\Desktop\project#2\Class#0'
os.chdir(root)

path_script = r'C:\Users\Wenxi\Desktop\project#2\Class#0\script'
if path_script not in sys.path:
    sys.path.append(path_script)
    
from func_absolute_value import *
   
os.chdir('.\input') 
files = os.listdir()

sub_file_list = []

for file in files:
    if file.lower().find('dep')>-1 and file.find('$')==-1:
        
        #get the department number
        dept_num = get_dept_from_filename(file)
        
        #locate the dataframe 
        dept_df_orig = pd.read_excel(file)
        
        #pre-process the dataframe to get the payment and budget
        dept_df = preprocess_df(dept_df_orig,
                                -1)
        #process to get the update the dataframe in the required format
        dept_df_updated = get_desired_df(dept_df,
                                         dept_num)
        
        #collect the processed dataframes by appending it to the list object        
        sub_file_list.append(dept_df_updated)

report_total = pd.concat(sub_file_list)
report_output = report_total.sort_values([*report_total.columns[:3]])

os.chdir(os.path.join(root, 'output')) 
report_output.to_excel(f'Report_{datetime.now().month}_{datetime.now().day}.xlsx', index=False)
