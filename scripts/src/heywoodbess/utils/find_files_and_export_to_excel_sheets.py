from __future__ import with_statement
import sys
import os
import json
import shutil
import logging

from typing import List, Union, Optional
import pandas as pd
from tqdm.auto import tqdm
from contextlib import contextmanager
from datetime import datetime

import traceback
from pathlib import Path



def find_files(directory,file_type):
    file_dirs = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(file_type):
                file_dirs.append(os.path.join(root, file))
    return file_dirs

def exportToExcelSheets(filename,sheetnames,arr_list_of_dicts):

    writer = pd.ExcelWriter(filename+'.xlsx',engine='xlsxwriter')   # Creating Excel Writer Object from Pandas  
    workbook=writer.book
    i = 0
    for list_of_dicts in arr_list_of_dicts:
       # worksheet=workbook.add_worksheet(sheetnames[i])
       # writer.sheets[sheetnames[i]] = worksheet
        df = pd.DataFrame(list_of_dicts)
        df.to_excel(writer,sheet_name=sheetnames[i],startrow=0 , startcol=0)   
        i=i+1
    writer.close()