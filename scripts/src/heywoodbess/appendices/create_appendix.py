import sys
import os
import json

import pandas as pd 
import numpy as np
import math
import glob

import ast
from pallet import load_specs_from_csv, load_specs_from_xlsx
from pallet import now_str
from typing import List, Optional, Callable
from pallet.specs import load_specs_from_multiple_xlsx
from pathlib import Path
from heywoodbess.appendices.appendix_generation_heywoodbess import generate_appendix_heywoodbess
from heywoodbess.appendices.replace_full_stops import replace_full_stops
from heywoodbess.utils.remove_underscores import caption_preprocessor

#For older versions of pallet (verisons before 3.15.1) - does not support loading in mutiple specs 
def create_appendix_heywoodbess_legacy(
                PLOT_RESULTS_DIR:str,
                APPENDIX_REPORT:str,
                DEFAULT_DIR:str,
                XLSX_PATH:str,
                SHEETS_TO_PROCESS,
                NOW_STR,
                x86:bool):
    
    print("Creating appendices ...")

    record_errors = []

    folders = []
    folders = [item for item in os.listdir(PLOT_RESULTS_DIR) if os.path.join(PLOT_RESULTS_DIR,item)]
    doc_number = 0

    for folder in folders:
        doc_number = doc_number + 1
        folder_path = os.path.join(PLOT_RESULTS_DIR,folder)

        replace_full_stops(folder_path)

        spec = spec = load_specs_from_xlsx(XLSX_PATH, sheet_names=SHEETS_TO_PROCESS)
        spec = spec[spec["Category"] == f"{folder}"]
        spec_copy = spec.copy()
        plot_path = os.path.join(PLOT_RESULTS_DIR,folder)


        for IS_CHARGING_PRESENT in [True,False]:
            if IS_CHARGING_PRESENT:
                try:
                    spec = spec[spec["File_Name"].str.contains("CRG", na=False)]
                    title = spec["Appendix_Name"].iloc[0]
                    if "Report" in spec.columns and spec["Report"].notna().all():
                        report_name = ast.literal_eval(spec_copy["Report"].iloc[0])
                        if len(report_name) > 1:
                            if x86:
                                report_psse = [item for item in report_name if 'psse' in item.lower()]
                                report = str(report_psse[0])

                            if not x86:
                                report_pscad = [item for item in report_name if 'pscad' in item.lower()]
                                report = str(report_pscad[0])
                        else:
                            report = str(report_name[0])   

                        APPENDIX_OUTPUT_DIR = os.path.join(APPENDIX_REPORT,report,f"Charging Appendix {NOW_STR}")

                    else:
                        APPENDIX_OUTPUT_DIR = os.path.join(DEFAULT_DIR,f"Charging Appendix {NOW_STR}")

                    output_path = os.path.join(APPENDIX_OUTPUT_DIR,f"{title}.pdf")
                    generate_appendix_heywoodbess(
                        project_name="heywoodbess",
                        client="Atmos",
                        title = title,
                        doc_number=f"GRID-LINK-DOC-NUMBER-{doc_number}",
                        issued_date="22 April 2025",
                        revision="Rev 0",
                        revision_history_csv_path="revisionhistory.csv",
                        plots_directory=plot_path,
                        output_path=output_path,
                        bess_charging=IS_CHARGING_PRESENT,
                        caption_preprocessor_fn=caption_preprocessor,
                        )
                except Exception as e:
                    generate_appendix_heywoodbess(
                        project_name="heywoodbess",
                        client="Atmos",
                        title = f"Appendix {folder} - Charging",
                        doc_number=f"GRID-LINK-DOC-NUMBER-{doc_number}",
                        issued_date="14 Feb 2025",
                        revision="Rev 0",
                        revision_history_csv_path="revisionhistory.csv",
                        plots_directory=plot_path,
                        output_path=os.path.join(DEFAULT_DIR,f"Charging Appendix {NOW_STR}",f"Appendix {folder}- Charging.pdf"),
                        bess_charging=IS_CHARGING_PRESENT,
                        caption_preprocessor_fn=caption_preprocessor,
                        )
                    print(f"Error: {e} --> NO 'Appendix_Name' column present in spec --> appendix stored in Default dir")
                    record_errors.append(f"{folder} - Charging Appendix: Stored in default dir. Error: {e}")

            else:
                try:
                    spec_copy = spec_copy[~spec_copy["File_Name"].str.contains("CRG", na=False)]
                    title = spec_copy["Appendix_Name"].iloc[0]
                    if "Report" in spec_copy.columns and spec_copy["Report"].notna().all():
                        report_name = ast.literal_eval(spec_copy["Report"].iloc[0])
                        if len(report_name) > 1:
                            if x86:
                                report_psse = [item for item in report_name if 'psse' in item.lower()]
                                report = str(report_psse[0])
                            if not x86:
                                report_pscad = [item for item in report_name if 'pscad' in item.lower()]
                                report = str(report_pscad[0])
                        else:
                            report = str(report_name[0])                     

                        APPENDIX_OUTPUT_DIR = os.path.join(APPENDIX_REPORT,report,f"Discharging Appendix {NOW_STR}")

                    else:
                        APPENDIX_OUTPUT_DIR = os.path.join(DEFAULT_DIR,f"Discharging Appendix {NOW_STR}")

                    output_path = os.path.join(APPENDIX_OUTPUT_DIR,f"{title}.pdf")
                    generate_appendix_heywoodbess(
                        project_name="heywoodbess",
                        client="Atmos",
                        title = title,
                        doc_number=f"GRID-LINK-DOC-NUMBER-{doc_number}",
                        issued_date="14 Feb 2025",
                        revision="Rev 0",
                        revision_history_csv_path="revisionhistory.csv",
                        plots_directory=plot_path,
                        output_path=output_path,
                        bess_charging=IS_CHARGING_PRESENT,
                        caption_preprocessor_fn=caption_preprocessor,
                        )
                except Exception as e:
                    generate_appendix_heywoodbess(
                        project_name="heywoodbess",
                        client="Atmos",
                        title = f"Appendix {folder}",
                        doc_number=f"GRID-LINK-DOC-NUMBER-{doc_number}",
                        issued_date="14 Feb 2025",
                        revision="Rev 0",
                        revision_history_csv_path="revisionhistory.csv",
                        plots_directory=plot_path,
                        output_path=os.path.join(DEFAULT_DIR,f"Discharging Appendix {NOW_STR}",f"Appendix {folder}.pdf"),
                        bess_charging=IS_CHARGING_PRESENT,
                        caption_preprocessor_fn=caption_preprocessor,
                        )
                    print(f"Error: {e} --> NO 'Appendix_Name' column present in spec --> appendix stored in Default dir")
                    record_errors.append(f"{folder} Appendix: Stored in default dir. Error: {e}")



    print("FINISHED APPENDIX GENERATION\n")
    print("All the Errors Encountered:")
    if len(record_errors)==0:
        print("No errors")
    else:
        for idx,error in enumerate(record_errors):
            print(f"{idx+1}. {error}\n")
    return

#Accomodates for new pallet version 3.15.1 - can load in multiple specs 
#This function will distinguish between CSR and other tests such as DMAT. This function will combine charging and discharging studies for CSR studies, and will separate discharging and charging appendices for DMAT studies. 
#Mulitiple plot directories can be provided, and it will create the appendices for all the studies provided in the plot_dir. 
def create_appendix_heywoodbess(
                PLOT_RESULTS_DIRS:list,
                OUTPUT_DIR_DMAT:str,
                OUTPUT_DIR_CSR,
                XLSX_PATHS:list,
                SHEETS_TO_PROCESS:list,
                NOW_STR,
                issued_date: str,
                revision_no: str,
                x86:bool,
                appendix_project_name: str):
    
    print("Creating appendices ...")

    record_errors = []
    for plot_dir in PLOT_RESULTS_DIRS:
        
        folders = [item for item in os.listdir(plot_dir) if os.path.join(plot_dir,item)]
        doc_number_CSR = 0
        doc_number_DMAT = 0

        for folder in folders:
            folder_path = os.path.join(plot_dir,folder)

            replace_full_stops(folder_path)

            spec = load_specs_from_multiple_xlsx(XLSX_PATHS, sheet_names=SHEETS_TO_PROCESS)

            spec = spec[spec["Category"] == f"{folder}"]
            spec_copy = spec.copy()
            spec_csr_copy = spec.copy()
            plot_path = os.path.join(plot_dir,folder)
            folder = str(folder)

            if "CSR" in folder:
                doc_number_CSR = doc_number_CSR + 1
                try:
                    if x86:
                        title = spec_csr_copy["Appendix_Name_PSSE"].iloc[0]
                    if not x86: 
                        title = spec_csr_copy["Appendix_Name_PSCAD"].iloc[0]
                    
                    APPENDIX_OUTPUT_DIR = os.path.join(OUTPUT_DIR_CSR,f"Appendix_{NOW_STR}")
                    output_path = os.path.join(APPENDIX_OUTPUT_DIR,f"{title}.pdf")


                    generate_appendix_heywoodbess(
                                    project_name=appendix_project_name,
                                    client="Atmos",
                                    title = title,
                                    doc_number=f"GRID-LINK-DOC-NUMBER-{doc_number_CSR}",
                                    issued_date=issued_date,
                                    revision=revision_no,
                                    revision_history_csv_path="revisionhistory.csv",
                                    plots_directory=plot_path,
                                    output_path=output_path,
                                    bess_charging=None,
                                    caption_preprocessor_fn=caption_preprocessor,
                                    )
                except Exception as e:
                    generate_appendix_heywoodbess(
                                project_name=appendix_project_name,
                                client="Atmos",
                                title = f"Appendix {folder}",
                                doc_number=f"GRID-LINK-DOC-NUMBER-{doc_number_CSR}",
                                issued_date=issued_date,
                                revision=revision_no,
                                revision_history_csv_path="revisionhistory.csv",
                                plots_directory=plot_path,
                                output_path=os.path.join(OUTPUT_DIR_CSR,f"Appendix_{NOW_STR}","Appendix {folder}.pdf"),
                                bess_charging=None,
                                caption_preprocessor_fn=caption_preprocessor,
                                )
                    print(f"Error: {e} --> NO 'Appendix_Name' column present in spec --> appendix stored in Default dir")
                    record_errors.append(f"{folder} Appendix: Stored in default dir. Error: {e}")

            else:
                doc_number_DMAT = doc_number_DMAT + 1
                for IS_CHARGING_PRESENT in [True,False]:
                    if IS_CHARGING_PRESENT:
                        try:
                            spec = spec[spec["File_Name"].str.contains("CRG", na=False)]
                            if x86:
                                title = spec["Appendix_Name_PSSE"].iloc[0]
                            if not x86:
                                title = spec["Appendix_Name_PSCAD"].iloc[0]
                        
                            APPENDIX_OUTPUT_DIR = os.path.join(OUTPUT_DIR_DMAT,f"Charging Appendix",NOW_STR)

                            output_path = os.path.join(APPENDIX_OUTPUT_DIR,f"{title}.pdf")

                            generate_appendix_heywoodbess(
                                project_name=appendix_project_name,
                                client="Atmos",
                                title = title,
                                doc_number=f"GRID-LINK-DOC-NUMBER-{doc_number_DMAT}",
                                issued_date=issued_date,
                                revision=revision_no,
                                revision_history_csv_path="revisionhistory.csv",
                                plots_directory=plot_path,
                                output_path=output_path,
                                bess_charging=IS_CHARGING_PRESENT,
                                caption_preprocessor_fn=caption_preprocessor,
                                )
                        except Exception as e:
                            generate_appendix_heywoodbess(
                                project_name=appendix_project_name,
                                client="Atmos",
                                title = f"Appendix {folder} - Charging",
                                doc_number=f"GRID-LINK-DOC-NUMBER-{doc_number_DMAT}",
                                issued_date=issued_date,
                                revision=revision_no,
                                revision_history_csv_path="revisionhistory.csv",
                                plots_directory=plot_path,
                                output_path=os.path.join(OUTPUT_DIR_DMAT,f"Charging Appendix",NOW_STR,f"Appendix {folder}.pdf"),
                                bess_charging=IS_CHARGING_PRESENT,
                                caption_preprocessor_fn=caption_preprocessor,
                                )
                            print(f"Error: {e} --> NO 'Appendix_Name_[Software]' column present in spec --> appendix stored in Default dir")
                            record_errors.append(f"{folder} - Charging Appendix: Stored in default dir. Error: {e}")

                    else:
                        try:
                            spec_copy = spec_copy[~spec_copy["File_Name"].str.contains("CRG", na=False)]
                            if x86:
                                title = spec_copy["Appendix_Name_PSSE"].iloc[0]
                            if not x86:
                                title = spec_copy["Appendix_Name_PSCAD"].iloc[0]
                                
                            APPENDIX_OUTPUT_DIR = os.path.join(OUTPUT_DIR_DMAT,f"Discharging Appendix", NOW_STR)

                            output_path = os.path.join(APPENDIX_OUTPUT_DIR,f"{title}.pdf")

                            generate_appendix_heywoodbess(
                                project_name=appendix_project_name,
                                client="Atmos",
                                title = title,
                                doc_number=f"GRID-LINK-DOC-NUMBER-{doc_number_DMAT}",
                                issued_date=issued_date,
                                revision=revision_no,
                                revision_history_csv_path="revisionhistory.csv",
                                plots_directory=plot_path,
                                output_path=output_path,
                                bess_charging=IS_CHARGING_PRESENT,
                                caption_preprocessor_fn=caption_preprocessor,
                                )
                        except Exception as e:
                            generate_appendix_heywoodbess(
                                project_name=appendix_project_name,
                                client="Atmos",
                                title = f"Appendix {folder}",
                                doc_number=f"GRID-LINK-DOC-NUMBER-{doc_number_DMAT}",
                                issued_date=issued_date,
                                revision=revision_no,
                                revision_history_csv_path="revisionhistory.csv",
                                plots_directory=plot_path,
                                output_path=os.path.join(OUTPUT_DIR_DMAT,f"Discharging Appendix",NOW_STR,f"Appendix {folder}.pdf"),
                                bess_charging=IS_CHARGING_PRESENT,
                                caption_preprocessor_fn=caption_preprocessor,
                                )
                            print(f"Error: {e} --> NO 'Appendix_Name' column present in spec --> appendix stored in Default dir")
                            record_errors.append(f"{folder} Appendix: Stored in default dir. Error: {e}")



    print("FINISHED APPENDIX GENERATION\n")
    print("All the Errors Encountered:")
    if len(record_errors)==0:
        print("No errors")
    else:
        for idx,error in enumerate(record_errors):
            print(f"{idx+1}. {error}\n")
    return



