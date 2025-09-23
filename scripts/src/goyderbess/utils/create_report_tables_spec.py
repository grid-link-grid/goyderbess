import sys
import os
import json
import pandas as pd 
import numpy as np
import math
import glob

import ast
from pallet import load_specs_from_csv, load_specs_from_xlsx
from pallet.specs import load_specs_from_multiple_xlsx
from pallet import now_str
from pathlib import Path
import traceback
def can_cast_to_float(series):
    return series.apply(lambda x: isinstance(x, (float, int))).all()

def manipulations(input_df : pd.DataFrame) -> pd.DataFrame:
    df = input_df.copy()

    # Fault imepdance
    df["Fault_Impedance"] = None
    for idx, row in df.iterrows():
        if pd.notna(row["Ures_sig"]):
            df.at[idx, "Fault_Impedance"] = f"{row['Ures_sig']} pu"
        elif pd.notna(row["Zf2Zs_sig"]):
            if pd.notna(row["Rf_Offset_sig"]) and row["Rf_Offset_sig"] > 0:
                df.at[idx, "Fault_Impedance"] = f"{row['Rf_Offset_sig']} Ω"
            else:
                df.at[idx, "Fault_Impedance"] = f"Zf = {row['Zf2Zs_sig']} Zs"

    return df

#This function will output the report tables for the clauses specified in the SHEETS_TO_PROCESS and column mapping csv. If there are no overlap between the two, no table csvs will be generated
def create_report_table(
                        COLMAP_PATH,
                        SPEC_PATH,
                        SHEETS_TO_PROCESS,
                        OUTPUT_DIR
                        ):

    print("Creating report tables ...")

    
    colmap_path = COLMAP_PATH
    spec_path = SPEC_PATH
    sheets_to_process = SHEETS_TO_PROCESS

    pd.reset_option('all')


    colmap_df = pd.read_csv(colmap_path)
    spec_df = load_specs_from_xlsx(spec_path, sheet_names=sheets_to_process, add_sheet_name_column=True)
    try:
        spec_df = manipulations(spec_df)

    except Exception as e:
        print("Error:",e)
        traceback.print_exc()
        print("Continuing to generate tables ...")
    

    spec_df["Sheet_Name"] = spec_df["Sheet_Name"].astype(str)
    for _, sheet_colmap in colmap_df.iterrows():
        sheet_colmap = sheet_colmap[sheet_colmap.notna()]
        
        if sheet_colmap['Sheet_To_Update'] in sheets_to_process: 
            sheet_name = str(sheet_colmap["Sheet_To_Update"])
            
            sheet_colmap = sheet_colmap.drop(labels=["Sheet_To_Update"])

            sheet_df = spec_df[spec_df["Sheet_Name"] == sheet_name]
            renaming_dict = dict(zip([str(x) for x in sheet_colmap.index.tolist()], [str(x) for x in sheet_colmap.values.tolist()]))
            sheet_df = sheet_df.rename(columns=renaming_dict)
            sheet_df = sheet_df[[str(x) for x in sheet_colmap.values.tolist()]]


            # Reference and results columns
            # sheet_df["Figure"] = [f"Figure {i+1}" for i in range(len(sheet_df))]
            # sheet_df.insert(0, "Figure", [f"Figure {i+1}" for i in range(len(sheet_df))])
            # sheet_df["Results"] = "Acceptable"

            # Cast columns to float if possible
            for col in sheet_df.columns:
                if can_cast_to_float(sheet_df[col]):
                    sheet_df[col] = sheet_df[col].astype(float)

            # Remove unbalanced faults for CSR
        
            if sheet_name == "324_325_Faults":
                for balanced in [True, False]:
                    if balanced:
                        sheet_df_copy = sheet_df.copy()
                        sheet_df_copy = sheet_df_copy[sheet_df_copy["Type"] == "3PHG"]
                        sheet_df_copy.to_csv(os.path.join(OUTPUT_DIR, f"{sheet_name}-balanced-test-table.csv".replace("_", "-")), index=False, float_format='%g')
                    else:
                        sheet_df_copy = sheet_df.copy()
                        sheet_df_copy = sheet_df_copy[sheet_df_copy["Type"] != "3PHG"]
                        sheet_df_copy.to_csv(os.path.join(OUTPUT_DIR, f"{sheet_name}-unbalanced-test-table.csv".replace("_", "-")), index=False, float_format='%g')
            
            sheet_df.to_csv(os.path.join(OUTPUT_DIR, f"{sheet_name}-test-table.csv".replace("_", "-")), index=False, float_format='%g')

    
    print("Finished table generation.")
    print("\n")
    return


#This function operates the same as create_report_table but is tailored for WAN
def create_report_table_WAN(
                            COLMAP_PATH,
                            SPEC_PATH,
                            SHEETS_TO_PROCESS,
                            OUTPUT_DIR
                            ):
    print("Creating report tables ...")
    

    colmap_path = COLMAP_PATH
    spec_path = SPEC_PATH
    sheets_to_process = SHEETS_TO_PROCESS

    pd.reset_option('all')

    # if os.path.exists(OUTPUT_DIR) and os.path.isdir(OUTPUT_DIR):
        # shutil.rmtree(OUTPUT_DIR)
    # os.mkdir(OUTPUT_DIR)

    colmap_df = pd.read_csv(colmap_path)
    spec_df = load_specs_from_xlsx(spec_path, sheet_names=sheets_to_process, add_sheet_name_column=True)
    try:
        spec_df = manipulations(spec_df)

    except Exception as e:
        print("Error:",e)
        print("Continuing to generate tables ...")
        pass

    spec_df["Sheet_Name"] = spec_df["Sheet_Name"].astype(str)


    for _, sheet_colmap in colmap_df.iterrows():
        sheet_colmap = sheet_colmap[sheet_colmap.notna()]
        if sheet_colmap['Sheet_To_Update'] in sheets_to_process: 
            sheet_name = str(sheet_colmap["Sheet_To_Update"])
            sheet_colmap = sheet_colmap.drop(labels=["Sheet_To_Update"])

            sheet_df = spec_df[spec_df["Sheet_Name"] == sheet_name]
            renaming_dict = dict(zip([str(x) for x in sheet_colmap.index.tolist()], [str(x) for x in sheet_colmap.values.tolist()]))
            sheet_df = sheet_df.rename(columns=renaming_dict)
            sheet_df = sheet_df[[str(x) for x in sheet_colmap.values.tolist()]]


            # Reference and results columns
            # sheet_df["Figure"] = [f"Figure {i+1}" for i in range(len(sheet_df))]
            # sheet_df.insert(0, "Figure", [f"Figure {i+1}" for i in range(len(sheet_df))])
            # sheet_df["Results"] = "Acceptable"

            # Cast columns to float if possible
            for col in sheet_df.columns:
                if can_cast_to_float(sheet_df[col]):
                    sheet_df[col] = sheet_df[col].astype(float)

            # Remove unbalanced faults for CSR
        
            if sheet_name == "324_325_Faults":
                for balanced in [True, False]:
                    if balanced:
                        sheet_df_copy = sheet_df.copy()
                        sheet_df_copy = sheet_df_copy[sheet_df_copy["Type"] == "3PHG"]
                        sheet_df_copy.to_csv(os.path.join(OUTPUT_DIR, f"wan-{sheet_name}-balanced-test-table.csv".replace("_", "-")), index=False, float_format='%g')

                    else:
                        sheet_df_copy = sheet_df.copy()
                        sheet_df_copy = sheet_df_copy[sheet_df_copy["Type"] != "3PHG"]
                        sheet_df_copy.to_csv(os.path.join(OUTPUT_DIR, f"wan-{sheet_name}-unbalanced-test-table.csv".replace("_", "-")), index=False, float_format='%g')

            

            sheet_df.to_csv(os.path.join(OUTPUT_DIR, f"wan-{sheet_name}-test-table.csv".replace("_", "-")), index=False, float_format='%g')

    print("Finished table generation.")
    print("\n")

    return


def create_report_table_benchmarking(
                        COLMAP_PATH,
                        SPEC_PATH,
                        SHEETS_TO_PROCESS,
                        OUTPUT_DIR
                        ):

    print("Creating report tables ...")


    
    colmap_path = COLMAP_PATH
    spec_path = SPEC_PATH
    sheets_to_process = SHEETS_TO_PROCESS

    pd.reset_option('all')


    colmap_df = pd.read_csv(colmap_path)
    spec_df = load_specs_from_xlsx(spec_path, sheet_names=sheets_to_process, add_sheet_name_column=True)
    try:
        spec_df = manipulations(spec_df)

    except Exception as e:
        print("Error:",e)
        traceback.print_exc()
        print("Continuing to generate tables ...")
    

    spec_df["Sheet_Name"] = spec_df["Sheet_Name"].astype(str)
    for _, sheet_colmap in colmap_df.iterrows():
        sheet_colmap = sheet_colmap[sheet_colmap.notna()]
        
        if sheet_colmap['Sheet_To_Update'] in sheets_to_process: 
            sheet_name = str(sheet_colmap["Sheet_To_Update"])
            
            sheet_colmap = sheet_colmap.drop(labels=["Sheet_To_Update"])

            sheet_df = spec_df[spec_df["Sheet_Name"] == sheet_name]
            renaming_dict = dict(zip([str(x) for x in sheet_colmap.index.tolist()], [str(x) for x in sheet_colmap.values.tolist()]))
            sheet_df = sheet_df.rename(columns=renaming_dict)
            sheet_df = sheet_df[[str(x) for x in sheet_colmap.values.tolist()]]


            # Reference and results columns
            # sheet_df["Figure"] = [f"Figure {i+1}" for i in range(len(sheet_df))]
            # sheet_df.insert(0, "Figure", [f"Figure {i+1}" for i in range(len(sheet_df))])
            sheet_df["Results"] = "Acceptable"

            # Cast columns to float if possible
            for col in sheet_df.columns:
                if can_cast_to_float(sheet_df[col]):
                    sheet_df[col] = sheet_df[col].astype(float)

            # Remove unbalanced faults for CSR
        
            if sheet_name == "324_325_Faults":
                for balanced in [True, False]:
                    if balanced:
                        sheet_df_copy = sheet_df.copy()
                        sheet_df_copy = sheet_df_copy[sheet_df_copy["Type"] == "3PHG"]
                        sheet_df_copy.to_csv(os.path.join(OUTPUT_DIR, f"{sheet_name}-balanced-test-table.csv".replace("_", "-")), index=False, float_format='%g')
                    # else:
                    #     sheet_df_copy = sheet_df.copy()
                    #     sheet_df_copy = sheet_df_copy[sheet_df_copy["Type"] != "3PHG"]
                    #     sheet_df_copy.to_csv(os.path.join(OUTPUT_DIR, f"{sheet_name}-unbalanced-test-table.csv".replace("_", "-")), index=False, float_format='%g')
            
            sheet_df.to_csv(os.path.join(OUTPUT_DIR, f"{sheet_name}-test-table.csv".replace("_", "-")), index=False, float_format='%g')

    
    print("Finished table generation.")
    print("\n")
    return



#This function will output all the sheets specified in the column mapping csv. This is different to the other function, that will only create csvs of the sheets that overlap in both sheets_to_process (list) and column mapping.csv
def create_report_table_all_sheets(
                        COLMAP_PATH,
                        SPEC_PATH,
                        SHEETS_TO_PROCESS,
                        OUTPUT_DIR
                        ):

    print("Creating report tables ...")

    
    colmap_path = COLMAP_PATH
    spec_path = SPEC_PATH
    sheets_to_process = SHEETS_TO_PROCESS

    pd.reset_option('all')


    colmap_df = pd.read_csv(colmap_path)
    spec_df = load_specs_from_xlsx(spec_path, sheet_names=sheets_to_process, add_sheet_name_column=True)
    try:
        spec_df = manipulations(spec_df)

    except Exception as e:
        print("Error:",e)
        traceback.print_exc()
        print("Continuing to generate tables ...")
    

    spec_df["Sheet_Name"] = spec_df["Sheet_Name"].astype(str)
    for _, sheet_colmap in colmap_df.iterrows():
        sheet_colmap = sheet_colmap[sheet_colmap.notna()]
        

        sheet_name = str(sheet_colmap["Sheet_To_Update"])
        
        sheet_colmap = sheet_colmap.drop(labels=["Sheet_To_Update"])

        sheet_df = spec_df[spec_df["Sheet_Name"] == sheet_name]
        renaming_dict = dict(zip([str(x) for x in sheet_colmap.index.tolist()], [str(x) for x in sheet_colmap.values.tolist()]))
        sheet_df = sheet_df.rename(columns=renaming_dict)
        sheet_df = sheet_df[[str(x) for x in sheet_colmap.values.tolist()]]


        # Reference and results columns
        # sheet_df["Figure"] = [f"Figure {i+1}" for i in range(len(sheet_df))]
        # sheet_df.insert(0, "Figure", [f"Figure {i+1}" for i in range(len(sheet_df))])
        # sheet_df["Results"] = "Acceptable"

        # Cast columns to float if possible
        for col in sheet_df.columns:
            if can_cast_to_float(sheet_df[col]):
                sheet_df[col] = sheet_df[col].astype(float)

        # Remove unbalanced faults for CSR
    
        if sheet_name == "324_325_Faults":
            for balanced in [True, False]:
                if balanced:
                    sheet_df_copy = sheet_df.copy()
                    sheet_df_copy = sheet_df_copy[sheet_df_copy["Type"] == "3PHG"]
                    sheet_df_copy.to_csv(os.path.join(OUTPUT_DIR, f"{sheet_name}-balanced-test-table.csv".replace("_", "-")), index=False, float_format='%g')
                else:
                    sheet_df_copy = sheet_df.copy()
                    sheet_df_copy = sheet_df_copy[sheet_df_copy["Type"] != "3PHG"]
                    sheet_df_copy.to_csv(os.path.join(OUTPUT_DIR, f"{sheet_name}-unbalanced-test-table.csv".replace("_", "-")), index=False, float_format='%g')
        
        sheet_df.to_csv(os.path.join(OUTPUT_DIR, f"{sheet_name}-test-table.csv".replace("_", "-")), index=False, float_format='%g')

    
    print("Finished table generation.")
    print("\n")
    return

#This function operates the same as create_report_table_all, but is tailored for WAN
def create_report_table_WAN_all_sheets(
                            COLMAP_PATH,
                            SPEC_PATH,
                            SHEETS_TO_PROCESS,
                            OUTPUT_DIR
                            ):
    print("Creating report tables ...")


    colmap_path = COLMAP_PATH
    spec_path = SPEC_PATH
    sheets_to_process = SHEETS_TO_PROCESS

    pd.reset_option('all')

    # if os.path.exists(OUTPUT_DIR) and os.path.isdir(OUTPUT_DIR):
        # shutil.rmtree(OUTPUT_DIR)
    # os.mkdir(OUTPUT_DIR)

    colmap_df = pd.read_csv(colmap_path)
    spec_df = load_specs_from_xlsx(spec_path, sheet_names=sheets_to_process, add_sheet_name_column=True)
    try:
        spec_df = manipulations(spec_df)

    except Exception as e:
        print("Error:",e)
        print("Continuing to generate tables ...")
        pass

    spec_df["Sheet_Name"] = spec_df["Sheet_Name"].astype(str)


    for _, sheet_colmap in colmap_df.iterrows():
        sheet_colmap = sheet_colmap[sheet_colmap.notna()]

        sheet_name = str(sheet_colmap["Sheet_To_Update"])
        sheet_colmap = sheet_colmap.drop(labels=["Sheet_To_Update"])

        sheet_df = spec_df[spec_df["Sheet_Name"] == sheet_name]
        renaming_dict = dict(zip([str(x) for x in sheet_colmap.index.tolist()], [str(x) for x in sheet_colmap.values.tolist()]))
        sheet_df = sheet_df.rename(columns=renaming_dict)
        sheet_df = sheet_df[[str(x) for x in sheet_colmap.values.tolist()]]


        # Reference and results columns
        # sheet_df["Figure"] = [f"Figure {i+1}" for i in range(len(sheet_df))]
        # sheet_df.insert(0, "Figure", [f"Figure {i+1}" for i in range(len(sheet_df))])
        # sheet_df["Results"] = "Acceptable"

        # Cast columns to float if possible
        for col in sheet_df.columns:
            if can_cast_to_float(sheet_df[col]):
                sheet_df[col] = sheet_df[col].astype(float)

        # Remove unbalanced faults for CSR
    
        if sheet_name == "324_325_Faults":
            for balanced in [True, False]:
                if balanced:
                    sheet_df_copy = sheet_df.copy()
                    sheet_df_copy = sheet_df_copy[sheet_df_copy["Type"] == "3PHG"]
                    sheet_df_copy.to_csv(os.path.join(OUTPUT_DIR, f"wan-{sheet_name}-balanced-test-table.csv".replace("_", "-")), index=False, float_format='%g')

                else:
                    sheet_df_copy = sheet_df.copy()
                    sheet_df_copy = sheet_df_copy[sheet_df_copy["Type"] != "3PHG"]
                    sheet_df_copy.to_csv(os.path.join(OUTPUT_DIR, f"wan-{sheet_name}-unbalanced-test-table.csv".replace("_", "-")), index=False, float_format='%g')

        
        
        sheet_df.to_csv(os.path.join(OUTPUT_DIR, f"wan-{sheet_name}-test-table.csv".replace("_", "-")), index=False, float_format='%g')

    print("Finished table generation.")
    print("\n")

    return

#Modified to take multiple specs, with new pallet version 3.15.1
def create_report_table_CSR_DMAT(
                        COLMAP_PATHS:list,
                        SPEC_PATHS:list,
                        SHEETS_TO_PROCESS:list,
                        OUTPUT_DIRS:list,
                        x86:bool
                        ):

    print("Creating report tables ...")

    for index,column_mapping_path in enumerate(COLMAP_PATHS):
        colmap_path = column_mapping_path

        pd.reset_option('all')

        colmap_df = pd.read_csv(colmap_path)
        spec_df = load_specs_from_multiple_xlsx(SPEC_PATHS,SHEETS_TO_PROCESS)
        try:
            spec_df = manipulations(spec_df)

        except Exception as e:
            print("Error:",e)
            traceback.print_exc()
            print("Continuing to generate tables ...")
        

        spec_df["Sheet_Name"] = spec_df["Sheet_Name"].astype(str)
        for _, sheet_colmap in colmap_df.iterrows():
            sheet_colmap = sheet_colmap[sheet_colmap.notna()]
            sheet_name = str(sheet_colmap["Sheet_To_Update"])

            sheet_colmap = sheet_colmap.drop(labels=["Sheet_To_Update"])
            sheet_df = spec_df[spec_df["Sheet_Name"] == sheet_name]
            if not x86:
                if "Appendix_Name_PSSE" in sheet_colmap.index:
                    sheet_colmap = sheet_colmap.drop(labels=["Appendix_Name_PSSE"])
            if x86:
                if "Appendix_Name_PSCAD" in sheet_colmap.index:
                    sheet_colmap = sheet_colmap.drop(labels=["Appendix_Name_PSCAD"]) 
            if sheet_df.empty:
                print(f"Error: No table content for {sheet_name}")
                print(f"Make sure {sheet_name} is present in sheet_to_process input")
                continue

            renaming_dict = dict(zip([str(x) for x in sheet_colmap.index.tolist()], [str(x) for x in sheet_colmap.values.tolist()]))
            sheet_df = sheet_df.rename(columns=renaming_dict)
            
            sheet_df = sheet_df[[str(x) for x in sheet_colmap.values.tolist()]]


            # Reference and results columns
            # sheet_df["Figure"] = [f"Figure {i+1}" for i in range(len(sheet_df))]
            # sheet_df.insert(0, "Figure", [f"Figure {i+1}" for i in range(len(sheet_df))])
            # sheet_df["Results"] = "Acceptable"

            # Cast columns to float if possible
            for col in sheet_df.columns:
                if can_cast_to_float(sheet_df[col]):
                    sheet_df[col] = sheet_df[col].astype(float)

            # Remove unbalanced faults for CSR
        
            if sheet_name == "324_325_Faults" or sheet_name == "324_325_Faults_CRG":
                for balanced in [True, False]:
                    if balanced:
                        sheet_df_copy = sheet_df.copy()
                        sheet_df_copy = sheet_df_copy[sheet_df_copy["Type"] == "3PHG"]
                        sheet_df_copy.to_csv(os.path.join(OUTPUT_DIRS[index], f"{sheet_name}-balanced-test-table.csv".replace("_", "-")), index=False, float_format='%g')
                    else:
                        sheet_df_copy = sheet_df.copy()
                        sheet_df_copy = sheet_df_copy[sheet_df_copy["Type"] != "3PHG"]
                        sheet_df_copy.to_csv(os.path.join(OUTPUT_DIRS[index], f"{sheet_name}-unbalanced-test-table.csv".replace("_", "-")), index=False, float_format='%g')
            
            else:
                sheet_df.to_csv(os.path.join(OUTPUT_DIRS[index], f"{sheet_name}-test-table.csv".replace("_", "-")), index=False, float_format='%g')


        
    print("Finished table generation.")
    print("\n")
    return