from __future__ import with_statement
import sys
import os
import json
import shutil
import logging

# MAIN_DIRECTORY = os.getcwd()
# sys_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSPY39'  #or where else you find the psspy.pyc
# sys.path.append(sys_path_PSSE)
# os_path_PSSE=r' C:\Program Files (x86)\PTI\PSSE34\PSSBIN'  # or where else you find the psse.exe
# os.environ['PATH'] += ';' + os.environ['PATH']

# PSSBIN_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSBIN"""
# PSSPY_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSPY39"""
# sys.path.append(PSSBIN_PATH)
# sys.path.append(PSSPY_PATH)
# import psse34
# import psspy
# from psspy import _i,_f,_s
from typing import List, Union, Optional
import pandas as pd
from tqdm.auto import tqdm
from contextlib import contextmanager
from datetime import datetime
# import redirect
import re
import glob
from pallet.specs import load_specs_from_xlsx
import traceback
from heywoodbess.utils.remove_underscores import caption_preprocessor
from heywoodbess.utils.find_files_and_export_to_excel_sheets import find_files
from heywoodbess.utils.find_files_and_export_to_excel_sheets import exportToExcelSheets

def runAnalysis_WAN_SS(inputsdir,outputsdir,filename):

    base_pretap_flows_df = pd.read_excel(os.path.join(inputsdir,r"""0_BASE CASE.xlsx"""),sheet_name='PRE-TAP FLOWS')
    base_posttap_flows_df = pd.read_excel(os.path.join(inputsdir,r"""0_BASE CASE.xlsx"""),sheet_name='POST-TAP FLOWS')
    base_pretap_volts_df = pd.read_excel(os.path.join(inputsdir,r"""0_BASE CASE.xlsx"""),sheet_name='PRE-TAP VOLTS')
    base_posttap_volts_df = pd.read_excel(os.path.join(inputsdir,r"""0_BASE CASE.xlsx"""),sheet_name='POST-TAP VOLTS')

    worksheets = find_files(inputsdir,".xlsx")
    def extract_number(filepath):
        filename = os.path.basename(filepath)
        try:
            number_str = filename.split('_')[0]
            return int(number_str)
        except ValueError:
            return float('inf') 
        
    worksheets_sorted = sorted(worksheets, key=extract_number)
  
    worksheets_sorted.remove(os.path.join(inputsdir,r"""0_BASE CASE.xlsx"""))

    change_pretap_flows_df = base_pretap_flows_df[['name', 'from','to','id']].copy() 
    change_posttap_flows_df = base_posttap_flows_df[['name', 'from','to','id']].copy()
    change_pretap_volts_df = base_pretap_volts_df[['name', 'number']].copy()
    change_posttap_volts_df = base_posttap_volts_df[['name', 'number']].copy()

    #Adding base case columns into dataframes so that they can be referenced in the output analysis
    change_pretap_flows_df[["Base Case"]] = base_pretap_flows_df[['PCTRTA (% RATING A)']]
    change_posttap_flows_df[["Base Case"]] = base_posttap_flows_df[['PCTRTA (% RATING A)']]
    change_pretap_volts_df[["Base Case"]] = base_pretap_volts_df[['PU']]
    change_posttap_volts_df[["Base Case"]] = base_posttap_volts_df[['PU']]

    base_pretap_volts_df_copy = base_pretap_volts_df.copy()
    base_posttap_volts_df_copy = base_posttap_volts_df.copy()
    base_pretap_flows_df_copy = base_pretap_flows_df.copy()
    base_posttap_flows_df_copy = base_posttap_flows_df.copy()

    base_pretap_flows_df_copy["Unique ID"] = base_pretap_flows_df_copy["from"].astype(str) + "_" + \
                                base_pretap_flows_df_copy["to"].astype(str) + "_" + \
                                base_pretap_flows_df_copy["id"].astype(str)

    base_posttap_flows_df_copy["Unique ID"] = base_posttap_flows_df_copy["from"].astype(str) + "_" + \
                                base_posttap_flows_df_copy["to"].astype(str) + "_" + \
                                base_posttap_flows_df_copy["id"].astype(str)


    case_name = str(os.path.basename(os.path.dirname(inputsdir)))
    for worksheet in worksheets_sorted:

        contingency_name = os.path.basename(worksheet).replace(".xlsx","")
        contingency_pretap_flows_df = pd.read_excel(worksheet,sheet_name='PRE-TAP FLOWS')
        contingency_posttap_flows_df = pd.read_excel(worksheet,sheet_name='POST-TAP FLOWS')
        contingency_pretap_volts_df = pd.read_excel(worksheet,sheet_name='PRE-TAP VOLTS')
        contingency_posttap_volts_df = pd.read_excel(worksheet,sheet_name='POST-TAP VOLTS')
    

        contingency_pretap_volts_df_copy = contingency_pretap_volts_df.copy()
        contingency_posttap_volts_df_copy = contingency_posttap_volts_df.copy()
        contingency_pretap_flows_df_copy = contingency_pretap_flows_df.copy()
        contingency_posttap_flows_df_copy = contingency_posttap_flows_df.copy()



        contingency_pretap_flows_df_copy["Unique ID"] = contingency_pretap_flows_df_copy["from"].astype(str) + "_" + \
                                      contingency_pretap_flows_df_copy["to"].astype(str) + "_" + \
                                      contingency_pretap_flows_df_copy["id"].astype(str)

        contingency_posttap_flows_df_copy["Unique ID"] = contingency_posttap_flows_df_copy["from"].astype(str) + "_" + \
                                      contingency_posttap_flows_df_copy["to"].astype(str) + "_" + \
                                      contingency_posttap_flows_df_copy["id"].astype(str)

        merged_pretap_volts = base_pretap_volts_df_copy.merge(contingency_pretap_volts_df_copy[["number","PU"]], on='number', suffixes=('_base','_cont'))
        merged_posttap_volts = base_posttap_volts_df_copy.merge(contingency_posttap_volts_df_copy[["number","PU"]], on='number', suffixes=('_base','_cont'))
        merged_pretap_flows = base_pretap_flows_df_copy.merge(contingency_pretap_flows_df_copy[["Unique ID", "PCTRTA (% RATING A)"]], on='Unique ID', suffixes=('_base','_cont'))
        merged_posttap_flows = base_posttap_flows_df_copy.merge(contingency_posttap_flows_df_copy[["Unique ID", "PCTRTA (% RATING A)"]], on='Unique ID', suffixes=('_base','_cont'))


        if "OOS" in case_name:
            conting_name = contingency_name+" OOS"
            change_pretap_flows_df[conting_name] = merged_pretap_flows['PCTRTA (% RATING A)_cont']
            change_posttap_flows_df[conting_name] = merged_posttap_flows['PCTRTA (% RATING A)_cont']
            change_pretap_volts_df[conting_name] = merged_pretap_volts['PU_cont'] - merged_pretap_volts['PU_base']
            change_posttap_volts_df[conting_name] = merged_posttap_volts['PU_cont'] - merged_posttap_volts['PU_base']
        if "IS" in case_name:
            conting_name = contingency_name+" IS"
            change_pretap_flows_df[conting_name] = merged_pretap_flows['PCTRTA (% RATING A)_cont']
            change_posttap_flows_df[conting_name] = merged_posttap_flows['PCTRTA (% RATING A)_cont']
            change_pretap_volts_df[conting_name] = merged_pretap_volts['PU_cont'] - merged_pretap_volts['PU_base']
            change_posttap_volts_df[conting_name] = merged_posttap_volts['PU_cont'] - merged_posttap_volts['PU_base']



    if not os.path.exists(outputsdir):
        os.makedirs(outputsdir)
    print()
    exportToExcelSheets(os.path.join(outputsdir,filename),["PRE-TAP FLOWS","PRE-TAP VOLTS DELTA","POST-TAP FLOWS","POST-TAP VOLTS DELTA"],[change_pretap_flows_df,change_pretap_volts_df,change_posttap_flows_df,change_posttap_volts_df])



def extract_interested_assets_WAN_SS(   #used to be called create_filtered_ss_analysis
        ASSETS_TO_PROCESS: list,
        ANALYSIS_XLSX_PATH:str,
        OUTPUT_DIR: str,
        COLMAP_PATH: str = None

):


    SHEETS_TO_PROCESS_ALL = [
        "PRE-TAP FLOWS",
        "POST-TAP FLOWS",
        "PRE-TAP VOLTS DELTA",
        "POST-TAP VOLTS DELTA"
]
    
    def can_cast_to_float(series):
        return series.apply(lambda x: isinstance(x, (float, int))).all()

    spec_path = os.path.join(ANALYSIS_XLSX_PATH)

    sheets_to_process = SHEETS_TO_PROCESS_ALL
    names_to_process = ASSETS_TO_PROCESS

    if "IS" in str(os.path.dirname(OUTPUT_DIR)):
        filename = "filtered_assets_of_interest_pre_post_IS"
    if "OOS" in str(os.path.dirname(OUTPUT_DIR)):
        filename = "filtered_assets_of_interest_pre_post_OOS"

    if COLMAP_PATH is not None:
        colmap_path = COLMAP_PATH
    else:
        print("No provided column mapping")
        sys.exit()

    pd.reset_option('all')

    colmap_df = pd.read_csv(colmap_path)
    spec_df = load_specs_from_xlsx(spec_path, sheet_names=sheets_to_process, add_sheet_name_column=True) 
    spec_df["name"] = spec_df["name"].astype(str)
    
    spec_df = spec_df[spec_df["name"].isin(names_to_process)]
    

    spec_df["Sheet_Name"] = spec_df["Sheet_Name"].astype(str)
    df_dict = {}

    for _, sheet_colmap in colmap_df.iterrows():
        sheet_colmap = sheet_colmap[sheet_colmap.notna()]
        sheet_name = str(sheet_colmap["Sheet_To_Update"])

        if sheet_name in sheets_to_process:   
 
            sheet_colmap = sheet_colmap.drop(labels=["Sheet_To_Update"])
            sheet_df = spec_df[spec_df["Sheet_Name"] == sheet_name]

            renaming_dict = dict(zip([str(x) for x in sheet_colmap.index.tolist()], [str(x) for x in sheet_colmap.values.tolist()]))
            
            sheet_df = sheet_df.rename(columns=renaming_dict)
            sheet_df = sheet_df[[str(x) for x in sheet_colmap.values.tolist()]]
            # Cast columns to float if possible
            for col in sheet_df.columns:
                if can_cast_to_float(sheet_df[col]):
                    sheet_df[col] = sheet_df[col].astype(float)
            sheet_df = sheet_df.reset_index(drop=True)
            df_dict[sheet_name] = sheet_df
  
    exportToExcelSheets(os.path.join(OUTPUT_DIR,filename),sheets_to_process,[df_dict[key] for key in sheets_to_process])

    print("Finished extracting interested assets --> exported as excel workbook.")
    print("\n")

    return

def remove_unwanted_cols(df_dict):
    for key,values in df_dict.items():
        df = df_dict[key]
        for df_column in df.columns:
            if df[df_column].isna().all():
                df = df.drop(columns=[df_column])
        df_dict[key] = df
    return df_dict

#This function will find scan through the WAN SS analysis excel and create a new excel containing all the found assets and contingencies that violated their given thresholds 
def find_violations_WAN_SS(
        SHEETS_TO_PROCESS:list,
        analysis_xlsx_path,
        col_map_contingency_path,  
        volt_thresold_pu:float,
        flow_threshold_percent:float
        ): 

    basename = os.path.basename(analysis_xlsx_path)
   
    if col_map_contingency_path is None:
        print("Did not provide contingency list")
        sys.exit()
    
    df = pd.read_csv(col_map_contingency_path, header=None)
    names_list = df.iloc[0].tolist()

    if "IS" in basename:
        suffix = "IS"
    if "OOS" in basename:
        suffix = "OOS"

    for idx, contingency in enumerate(names_list):
        if "IS" in basename:
            names_list[idx] = contingency + "" + suffix
        if "OOS" in basename:
            names_list[idx] = contingency + "" + suffix
        

    df_dict_violations={}
    for sheet in SHEETS_TO_PROCESS:
        
        spec_df = load_specs_from_xlsx(analysis_xlsx_path, sheet_names=[sheet], add_sheet_name_column=True)   

        for idx, row in spec_df.iterrows():
            if "VOLT" in sheet:
                for column in names_list:
                    if abs(row[column]) <= volt_thresold_pu:
                        spec_df.at[idx,column] = None
                spec_df = spec_df[~(spec_df[names_list].isna().all(axis=1))]
        
            if "FLOW" in sheet:
                for column in names_list:
                    if abs(row[column]) <= flow_threshold_percent:
                        spec_df.at[idx,column] = None
                spec_df = spec_df[~(spec_df[names_list].isna().all(axis=1))]
                

        spec_df = spec_df.rename(columns={"name":"ASSET"})
        spec_df = spec_df.drop(columns=["Sheet_Name"])
        spec_df = spec_df.drop(columns=["Unnamed: 0"])

        if "FLOW" in sheet:
            spec_df = spec_df.drop(columns=["from","to"])

        for column_name in spec_df.columns:
            new_name = caption_preprocessor(caption=column_name)
            spec_df = spec_df.rename(columns={column_name:new_name})

        df_dict_violations[sheet] = spec_df

    remove_unwanted_cols(df_dict_violations) #Will keep only the relevant columns in the dataframe

    print("Done finding violations.")
    return df_dict_violations



#This code will provide an excel workbooks containing, assets and contingencies that violated the thresholds, comparison excel workbook used to compare the voltages/flows for the same asset contingency combination found in the violation excel.
#It will also provide the difference between the violated values (volt and flows) and the comparison values for the given asset and contingency. 
def run_exacerbate_analysis(
        df_dict_violations,
        comparison_xlsx_path
                            ):
    if not os.path.exists(comparison_xlsx_path):
        print("Provided invalid comparison excel path. Aborting.")
        sys.exit()

    df_dict_comparison = {} #This will be the dictionary with all the comparisons to make with the violated assests in either the IS or OOS 
    df_dict_difference = {}
    for key,value in df_dict_violations.items():
        df_comp = load_specs_from_xlsx(comparison_xlsx_path, sheet_names=[key], add_sheet_name_column=True)
        df_v = df_dict_violations[key]

        if not df_v.empty:
            df_comp = df_comp.rename(columns={"name":"ASSET"})

            for cols_comp in df_comp.columns: #Setting columns in comparison df to be the same as the violation df
                new_name = caption_preprocessor(caption=cols_comp) #gets rid of any underscores in column names 
                df_comp = df_comp.rename(columns={cols_comp:new_name})
                if "IS" in cols_comp:
                    df_comp = df_comp.rename(columns={new_name:new_name.replace("IS","")})

                if "OOS" in cols_comp:
                    df_comp = df_comp.rename(columns={new_name:new_name.replace("OOS","")})

            for cols_v in df_v.columns:
                new_name = caption_preprocessor(caption=cols_v)
                df_v = df_v.rename(columns={cols_v:new_name})
                if "IS" in cols_v:
                    df_v = df_v.rename(columns={new_name:new_name.replace("IS","")})        
                if "OOS" in cols_v:
                    df_v = df_v.rename(columns={new_name:new_name.replace("OOS","")})   

            common_contingencies = [common_cols for common_cols in df_v.columns if common_cols in df_comp.columns]

            df_v = df_v[common_contingencies]
            df_comp = df_comp[common_contingencies]

            if "VOLT" in key:
                df_comp = df_comp[df_comp["number"].isin(df_v["number"])]
            if "FLOW" in key:
                df_comp = df_comp[df_comp["ASSET"].isin(df_v["ASSET"])]
                df_comp = df_comp[df_comp["id"].isin(df_v["id"])]

            
            df_v.set_index("ASSET", inplace=True)
            df_comp.set_index("ASSET", inplace=True)
           
            df_comp = df_comp.reindex(df_v.index) #Making sure the order of the indexes (assets) for df comparison is the same as the violation in order to compare
    
        
            for col in df_v.columns:

                df_comp[col] = df_v[col].where(df_v[col].isna(),df_comp[col]) #df violation might have some empty rows that are irrelvant, the comparison df should mimic this behaviour and keep only the interested values for comparison

            df_dict_comparison[key] = df_comp
            df_dict_violations[key] = df_v

            df_comp_abs = df_comp.abs()
            df_v_abs = df_v.abs()

            df_diff = df_v_abs.subtract(df_comp_abs)
            df_diff = df_diff.where(df_v.notna(),None)
            if "Base Case" in df_diff.columns:
                df_diff = df_diff.drop(columns=["Base Case"])
            
            df_dict_difference[key] = df_diff
        else:
            df_dict_comparison[key] = df_v
            df_dict_violations[key] = df_v
            df_dict_difference[key] = df_v
    return df_dict_comparison,df_dict_violations,df_dict_difference
