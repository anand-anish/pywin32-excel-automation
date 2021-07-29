import win32com.client as win32
import pandas as pd
import numpy as np
from pathlib import Path
import re
import sys
import os
import json

# Open excel workbook based on path and name of workbook
def run_excel(f_path, f_name):

    file = f_path +'/'+ f_name

    # create excel object
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    # excel can be visible or not
    excel.Visible = True  # False
    
    # try except for file / path
    try:
        wb = excel.Workbooks.Open(file)
    except com_error as e:
        if e.excepinfo[5] == -2146827284:
            print(f'Failed to open spreadsheet.  Invalid filename or location: {file}')
        else:
            raise e
        sys.exit(1)
        
    return wb

# Describe the structure of pivot table based on name - row fields, column fields, filters etc.
 def pivot_description(pvtTable,selection=None):
    key = [str(j) for i,j in enumerate(pvtTable.PageRange) if i%2 == 0]
    value = [str(j) for i,j in enumerate(pvtTable.PageRange) if i%2 == 1]

    pvtFilter = dict(zip(key, value))

    print('Filters : \n')
    for fltr_nm,fltr_val in pvtFilter.items():
        print("{} ({})".format(fltr_nm,fltr_val)) 

    row_fields_item = []
    print('\nRow Fields : \n')
    for i in pvtTable.GetRowFields():
        row_fields_item.append(str(i))
        print(str(i))

    print('\nColumn Fields : \n')
    column_fields_item = []
    for i in pvtTable.GetColumnFields():
        print(str(i))
        column_fields_item.append(str(i))

    print('\nSelected Column Metrics : \n')
    for i in pvtTable.GetColumnFields():
        for j in pvtTable.PivotFields(str(i)).VisibleItemsList: # selected filtered list
            print(j)    
    
    print('\nData Fields : \n')
    for i in pvtTable.GetDataFields():
        print(str(i))
    
    return
  
# Update Filter of Pivot table based on filter name and value 
 def pivot_update_filtr(pvtTable, filtr_nm, filtr_val):
    
    try:
        selected_filtr = ''.join([str(i) for i in pvtTable.PivotFields() if ("["+str(filtr_nm)+"]") in str(i)])
        #print('Selected Filter: ' ,selected_filtr)
        print("Current Filter Value : ",pvtTable.PivotFields(selected_filtr).CurrentPageName)
        tmp = str.split(selected_filtr,'.')
        tmp[len(tmp)-1] = "&["+str(filtr_val)+"]"
        val = '.'.join(tmp)
        print('Applied Filter : ',val)

        # Update Pivot Table
        pvtTable.PivotFields(selected_filtr).ClearAllFilters()
        pvtTable.PivotFields(selected_filtr).CurrentPageName = val
    except:
        print('\nError : Specified Filter Name is not present in this pivot table')

    return
  
  
# Expand/Collapse the rows of pivot table (for easy reading into the python environment)

def pivot_table_expand_collapse_row(pvtTable,row_fields,expand=True,repeatLables=True):
    if row_fields.lower()=='all':
        col =[]
        for i in pvtTable.RowFields:
            tmp = str.split(str(i),'.')
            tmp = tmp[len(tmp)-1]
            col.append(tmp.replace("[","").replace("]",""))
    else:
        col = row_fields    
    # expand/collapse based on setting expand parameter as true/false
    for i in col:
        col_nm = [str(j) for j in pvtTable.RowFields if ("["+str(i)+"]") in str(j)]
        col_nm = ''.join(col_nm)
        pvtTable.PivotFields(col_nm).DrilledDown = True
        print("Expanded : ", col_nm)
    
    # Repeat All labels in hierarchy
    if repeatLables is True:
        pvtTable.RepeatAllLabels(2)    # 1 = xlDoNotRepeatLabels ; 2 = xlRepeatLabels
        print("\nSuccess : Row Labels Repeated")
    else:
        pvtTable.RepeatAllLabels(1)
        print("\nSuccess : Row Labels Not Repeated")
        
  # Get pivot data based on table name and number of columns. Since this reads the entire pivot are, hence you might have to perform slight data processing to get the required field names
  def get_pivot_data(pvtTable,num_fields):
    
    # Extract Table Data
    table_data = []
    for i in pvtTable.TableRange1:
        table_data.append(str(i))
    
    # number of rows and columns
    num_row_fields = len(pvtTable.GetRowFields())
    num_col_fields = len([str(j) for i in pvtTable.GetColumnFields() for j in pvtTable.PivotFields(str(i)).VisibleItemsList])
    tot_cols = int(num_row_fields + num_col_fields)
    tot_rows = int(len(table_data)/num_fields)
    
    # reshape the table based on rows and columns
    arr2D = np.reshape(table_data, (tot_rows,num_fields))
    df = pd.DataFrame(arr2D)
    return(df)
  
  
