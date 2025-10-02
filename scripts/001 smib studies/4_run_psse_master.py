import sys
import os
import shutil
from pallet.utils.time_utils import now_str
import numpy as np
from pathlib import Path


import json
import os, glob
import pandas as pd

from pallet.specs import load_specs_from_csv, load_specs_from_xlsx
from pallet.config import get_config

from gridlink.utils.wan.wan_utils import find_files

from goyderbess.plotting.goyderbessPSSEPlotter import goyderbessPssePlotter
from goyderbess.plotting.process_and_calc_psse import pre_process_dataframe
from goyderbess.studyrunners.psse_study_runner import run_psse_studies
from goyderbess.analysis.run_analysis_psse import run_analysis_psse
from goyderbess.appendices import create_appendix

from goyderbess.plotting.replotters import replot_psse
from goyderbess.appendices.create_appendix import create_appendix_goyderbess
from goyderbess.utils.create_report_tables_spec import create_report_table_CSR_DMAT

from goyderbess.studyrunners.psse_study_runner import get_vslacks

from pallet.specs import load_specs_from_multiple_xlsx
#------------------------------------------------------------------ CLAUSES --------------------------------------------------------------------------------------
x86 = True
NOW_STR = now_str()


SHEETS_TO_PROCESS_CSR = [
    # "5253_F",
    # "5254_Withstand",
    # "5254_CUO",
    # "5255_BalFaults",
    # "5255_UnbalFaults",
    # "52511_Fgrid",
    # "52513_Vref",
    # "52513_Qref",
    # "52513_PFref",
    # "52514_Pref",
    # "5255_TOV",
    # "52513_Vgrid_droop",
]


SHEETS_TO_PROCESS_DMAT = [
    "DMATFLAT",
    # "324_325_Faults",
    # "326_MFRT", 
    # "329_TOV", 
    # "3210_Vref", 
    # "3210_Qref", 
    # "3210_PFref",
    # "3211_Pref",  #working
    # "3217_Pref_POC_SCR1",
    # "3212_Fgrid",
    # "3210_3214_Vgrid",
    # "3216_PhaseSteps",
    # "3218_SCR_Change_Faults_SCR1",
    # "SCR_Change_Faults_POC",

    # # "3220_Input_Power",
    # # "Initialisation_2"
]

SHEETS_TO_PROCESS_DMAT_CRG = [
 
    # "DMATFLAT_CRG",
    # "324_325_Faults_CRG",
    # "326_MFRT_CRG",
    # "329_TOV_CRG",
    # "3210_Vref_CRG",
    # "3210_Qref_CRG",
    # "3210_PFref_CRG",
    # "3211_Pref_CRG",
    # "3217_Pref_POC_SCR1_CRG",
    # "3212_Fgrid_CRG",
    # "3210_3214_Vgrid_CRG",
    # "3216_PhaseSteps_CRG",
    # "3218_SCR_Change_Faults_SCR1_CRG",
    # "SCR_Change_Faults_POC_CRG",

    # "Init_CRG"
]

SHEETS_TO_PROCESS_FLATRUN = [
    # "CORNER_POINTS",
]


ANALYSIS_TO_RUN = [

            # "dP/df characteristic",           
            
             
            # "CUO",                            

# # 5.2.5.13
            # "Vdroop characteristic",
            # "Vgrid step analysis",       
            # "Vref step analysis",              
            # "Qref step analysis",                
            # "PFref step analysis",
          
# # # # 5.2.5.8
            # "S5258 Active Power Reduction"        
# # # 5.5
            # "diq/dV characteristic",   
            #  "IQ Rise Settle & P Recovery Curve",               
# # # 5.3
#             "Frequency ride-through characteristic",
# # # 5.4
#             "Voltage ride-through characteristic",

                     
            ]

#------------------------------------------------------------------ END OF CLAUSES --------------------------------------------------------------------------------------



#----------------------------------------------------------------------------- SET RUN SETTINGS + PATHS  ------------------------------------------------------------------------------
SAV_VERSION = "v0-0-4"
DYR_VERSION = "v0-0-4"

# XLSX_DIR = r"C:\Users\lhyett.GRID-LINK\Grid-Link\Projects - PROJECT_SPEC (1)"
# XLSX_PATH_CSR = os.path.join(XLSX_DIR,"HY_Spec_CSR_300.xlsx")
# XLSX_PATH_DMAT = os.path.join(XLSX_DIR,"HY_Spec_DMAT.xlsx")
# XLSX_PATH_DMAT_CRG = os.path.join(XLSX_DIR,"HY_Spec_DMAT_CRG.xlsx")
# XLSX_PATH_DMAT_FLATRUN = os.path.join(XLSX_DIR,"HY_Spec_FLATTEST.xlsx")
#FLATRUN_XLSX = os.path.join(r"C:\Grid\chen\cg\psse\flatrun_spec","CGBess_Spec_Flatrun.xlsx")

XLSX_DIR = get_config("goyderbess", "spec_path")
XLSX_PATH_CSR = os.path.join(XLSX_DIR,"GO_Spec_CSR_300.xlsx")
XLSX_PATH_DMAT = os.path.join(XLSX_DIR,"GO_Spec_DMAT.xlsx")
XLSX_PATH_DMAT_CRG = os.path.join(XLSX_DIR,"GO_Spec_DMAT_CRG.xlsx")
XLSX_PATH_TESTING = os.path.join(XLSX_DIR,"GO_Spec_FLATTEST.xlsx")
PROJECT_DIR = get_config("goyderbess", "project_directory")

RUN_STUDIES = True
if RUN_STUDIES:
    slack_bus_num = 500001
    MODEL_DIR = os.path.join(PROJECT_DIR, r"""PSSE""")
    plotter = goyderbessPssePlotter(pre_process_fn=pre_process_dataframe)
    RESULTS_DIR = os.path.join(r'''C:\Grid\Goyder BESS results\PSSE''',f"{SAV_VERSION}_sav_{DYR_VERSION}_dyr", now_str())
    if not os.path.exists(RESULTS_DIR):
        os.makedirs(RESULTS_DIR)

    #Filtering spec
    specification = load_specs_from_multiple_xlsx([XLSX_PATH_CSR,XLSX_PATH_DMAT,XLSX_PATH_DMAT_CRG], sheet_names=[SHEETS_TO_PROCESS_CSR,
                                                                                                                               SHEETS_TO_PROCESS_DMAT,
                                                                                                                               SHEETS_TO_PROCESS_DMAT_CRG])
    spec = get_vslacks(specification,MODEL_DIR,slack_bus_num)
    if "Vslack_pu_psse" in spec.columns:
        spec_non_vslack = specification[~specification["Vslack_pu_psse"].astype(str).str.contains(r"\$VSLACK")]
    else:
        spec_non_vslack = specification 
    spec = pd.concat([spec, spec_non_vslack], ignore_index=True)
    spec = spec[spec['PSSE'] == True]
    # print(spec["Vslack_pu_psse"])
    # spec_vslack_df = pd.DataFrame(spec, columns=["Vslack_pu_psse"])
    # spec_vslack_df.to_csv("spec.csv")
    # sys.exit()
    

    # spec = spec[spec['Test No'] == 3]
    # spec = spec[spec['Test No'] <=167]
    #print(PREF_MW)
    # print(spec)0
    # spec = spec[spec['Batch'] == 2]

RUN_ANALYSIS = False
if RUN_ANALYSIS:
    if RUN_STUDIES:
        CSR_INPUTS_DIR= RESULTS_DIR
    else:
        CSR_INPUTS_DIR = r"""D:\alvin\Heywood work folder\analysis_testresult\results\psse\v0-2-0_sav_v0-2-0_dyr\20250522_1339_51718295"""
    analysis_extension = ".out" 
    ANALYSIS_OUTPUTS_DIR = os.path.join(r'''D:\alvin\results\goyderbess\analysis''',f"_analysis_{NOW_STR}")

    DPDF_CHARACTERISTIC_POINTS = [(-9.94, 600), (-4.97, 300), (-0.015, 0), (0.015, 0), (4.97, -300), (9.94, -600)] #Updated for  Heywood
    VDROOP_CHARACTERISTIC_POINTS = [(-0.1, 118.5),(-0.02, 118.5), (0.02, -118.5), (0.1, -118.5)] # Updated for Heywood

    HVRT_THRESHOLDS = [
            {
                "value": 1.2,
                "withstand_sec": 3,
                "colour": "orange",
            },
            {
                "value": 1.15,
                "withstand_sec": 20,
                "colour": "yellow",
            },
        ]
    LVRT_THRESHOLDS = [
            {
                "value": 0.8,
                "withstand_sec": 5,
                "colour": "red",
            }
        ]


#Replot PSSE studies 
REPLOT_PSSE = False
if REPLOT_PSSE:
    extension_replot = ".out"
    replotter = goyderbessPssePlotter(pre_process_fn=pre_process_dataframe)                                        
    PLOT_INPUTS_DIR = r"""C:\Grid\chen\cg\results\psse\v1-1-0_sav_v1-1-0_dyr\20250326_1234_16856156"""
    PLOT_OUT_PATH = os.path.join(r"""C:\Grid\chen\cg\results\psse\replots""") #DIR where you want png and pdf outputs of replots to be located


#Generate appendices for chosen clauses 
CREATE_APPENDIX = False
if CREATE_APPENDIX:
    PLOT_RESULTS_DIRS = [r"C:\Grid\cg\psse\example_outfiles_lmth"] #list of paths, can point to multiple directories that contains the plots.
    OUTPUT_DIR_DMAT = os.path.join(r"C:\Grid\cg\pscad\appendices\appendix_dmat_psse") # DEFAULT DIR where appendices should be stored if you do not have column in spec called "Report". You should set it as a default 
    OUTPUT_DIR_CSR = os.path.join(r"C:\Grid\cg\pscad\appendices\appendix_csr_psse")
    SHEETS_TO_PROCESS_COMBINED = [SHEETS_TO_PROCESS_CSR,SHEETS_TO_PROCESS_DMAT,SHEETS_TO_PROCESS_DMAT_CRG]
    XLSX_PATHS = [XLSX_PATH_CSR,XLSX_PATH_DMAT,XLSX_PATH_DMAT_CRG]
    issued_date = "22nd of April"
    revision_no = "rev 0"

    if not os.path.exists(OUTPUT_DIR_DMAT):
        os.makedirs(OUTPUT_DIR_DMAT)
    if not os.path.exists(OUTPUT_DIR_CSR):
        os.makedirs(OUTPUT_DIR_CSR)


#Create report tables for chosen clauses 
CREATE_REPORT_TABLES = False 
if CREATE_REPORT_TABLES:
    SHEETS_TO_PROCESS_TABLES = [SHEETS_TO_PROCESS_CSR,SHEETS_TO_PROCESS_DMAT,SHEETS_TO_PROCESS_DMAT_CRG]
    SPEC_PATHS_TABLES = [XLSX_PATH_CSR,XLSX_PATH_DMAT,XLSX_PATH_DMAT_CRG]
    TABLE_DIR = r"""C:\Grid\cgbess\scripts\005_generate_report_tables"""
    COLMAP_PATHS = [os.path.join(TABLE_DIR, "column_mapping_DMAT.csv"),os.path.join(TABLE_DIR, "column_mapping_CSR.csv")]
    OUTPUT_DIRS_TABLE = [r"""C:\Grid\cg\pscad\tableoutputsDMAT""",r"""C:\Grid\cg\pscad\tableoutputsCSR"""]
    for output_dirs in OUTPUT_DIRS_TABLE:
        if not os.path.exists(output_dirs):
            os.makedirs(output_dirs)




#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
if __name__ == "__main__":

    if RUN_STUDIES:
        run_psse_studies(
            spec=spec,
            plotter=plotter,
            MODEL_DIR=MODEL_DIR,
            RESULTS_DIR=RESULTS_DIR
        )


    if RUN_ANALYSIS:
        run_analysis_psse(
            x86=x86,
            extension=analysis_extension,
            ANALYSIS_TO_RUN=ANALYSIS_TO_RUN,
            DPDF_CHARACTERISTIC_POINTS=DPDF_CHARACTERISTIC_POINTS,
            VDROOP_CHARACTERISTIC_POINTS=VDROOP_CHARACTERISTIC_POINTS,
            HVRT_THRESHOLDS=HVRT_THRESHOLDS,
            LVRT_THRESHOLDS=LVRT_THRESHOLDS,
            CSR_INPUTS_DIR=CSR_INPUTS_DIR,
            OUTPUTS_DIR=ANALYSIS_OUTPUTS_DIR,
        )


    if REPLOT_PSSE:
        replot_psse(
            extension=extension_replot,
            replotter=replotter,
            PLOT_INPUTS_DIR=PLOT_INPUTS_DIR,
            PLOT_OUT_DIR=PLOT_OUT_PATH,
            x86=x86)


    if CREATE_APPENDIX:
        create_appendix(
                PLOT_RESULTS_DIRS=PLOT_RESULTS_DIRS,
                OUTPUT_DIR_DMAT=OUTPUT_DIR_DMAT,
                OUTPUT_DIR_CSR=OUTPUT_DIR_CSR,
                XLSX_PATHS=XLSX_PATHS,
                SHEETS_TO_PROCESS=SHEETS_TO_PROCESS_COMBINED,
                NOW_STR=NOW_STR,
                issued_date=issued_date,
                revision_no=revision_no,
                x86=x86)



    if CREATE_REPORT_TABLES:
        create_report_table_CSR_DMAT(
            COLMAP_PATHS=COLMAP_PATHS,
            SPEC_PATHS=SPEC_PATHS_TABLES,
            SHEETS_TO_PROCESS=SHEETS_TO_PROCESS_TABLES,
            OUTPUT_DIRS=OUTPUT_DIRS_TABLE,
            x86=x86)