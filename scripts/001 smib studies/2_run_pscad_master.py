import sys
import os
import json

import pandas as pd 
import numpy as np

from pallet import load_specs_from_csv, load_specs_from_xlsx
from pallet.specs import load_specs_from_multiple_xlsx
from pallet.utils.time_utils import now_str
from pallet import now_str
from pallet.pandapower.initialisation import calc_tov_shunt_mvar_from_ppoc_qpoc_vpoc_vslack
from pallet.config import get_config
from pathlib import Path

from goyderbess.plotting.replotters import replot_pscad
from goyderbess.appendices.create_appendix import create_appendix_goyderbess
from goyderbess.utils.create_report_tables_spec import create_report_table,create_report_table_CSR_DMAT
from goyderbess.studyrunners.pscad_study_runner import run_pscad_studies
from goyderbess.studyrunners.temp_pscad_study_runner import run_pscad_studies_temp
from goyderbess.analysis.run_analysis_pscad import run_analysis_pscad
from goyderbess.plotting.goyderbessPscadPlotter import goyderbessPscadPlotter
from goyderbess.plotting.goyderbessPscadReplotter import goyderbessPscadReplotter



#----------------------------------------------------------------------------- CLAUSES TO RUN --------------------------------------------------------------------------------------
NOW_STR = now_str()

x86 = False

APPENDIX_PROJECT_NAME = "Goyder BESS"

SHEETS_TO_PROCESS_CSR = [
    # "5253_F", #working
    # "5254_Withstand", #bad results
    # "5254_Withstand_test"
    # "5254_CUO", #working
    # "5255_TOV", #working
    # "5255_UnbalFaults",
    # "5255_BalFaults",
    # "5257_PLR",
    # "52511_Fgrid",
    # "52513_Vgrid_droop",
    # "52513_Vref",
    # "52513_Qref",
    # "52513_PFref",
    # "52514_Pref",
    # "5258_Fprotection",
    # "5258_Vprotection",
    # "52515_SCR_Withstand_Faults",
    # "52515_Vgrid",
    # "52515_SCR_Change_NoFault",
    # "5258_ActivePowerReduction"
]


SHEETS_TO_PROCESS_DMAT = [
    # "DMATFLAT",
    # "324_325_Faults",
    # "326_MFRT",
    # "329_TOV",
    # "3210_Vref",
    # "3210_Qref",
    "3210_PFref",
    # "3211_Pref",
    # "3217_Pref_POC_SCR1",
    # "3212_Fgrid",
    # "3210_3214_Vgrid",
    # "3215_ORT",
    # "3216_PhaseSteps",
    # "3218_SCR_Change_Faults_SCR1",
    # "SCR_Change_Faults_POC",
    # "3220_Input_Power",
    # "Initialisation_2",
    # "Init"

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
    # "3215_ORT_CRG",
    # "3216_PhaseSteps_CRG",
    # "3218_SCR_Change_Faults_SCR1_CRG",
    # "SCR_Change_Faults_POC_CRG",
]

SHEETS_TO_PROCESS_FLAT = [

    # "CORNER_POINTS",
    # "DMAT_POINTS",
    # "LOW_SCR_POINTS",
]




ANALYSIS_TO_RUN = [
    # "dP/df characteristic", #working
    # "CUO", #working
    # "Vdroop characteristic", #working
    # "Vgrid step analysis", #working
    # "Vref step analysis", #working
    # "diq/dV characteristic", #partially working - need to check on per unit reporting of Iq and Id at PoC
    # "Max fault current", #not working
    # "Frequency ride-through characteristic", #working
    # "Voltage ride-through characteristic", #working
    # "MFRT", #not working until you find inverter operation key INV_OPERATION_STATE_SPEC_KEY INV_ERR_MSG_SPEC_KEY
    # "IQ Rise Settle & P Recovery Curve", #working
    # "Phase Health Analysis", #working
    # "ORT Analysis", #working
    # "S5258 Active Power Reduction", #working
    # "PFref step analysis", #working
    # "Qref step analysis" #working

    ]

    
#----------------------------------------------------------------------------- END OF CLAUSES TO RUN ----------------------------------------------------------------------------------




#----------------------------------------------------------------------------- SET RUN SETTINGS + PATHS  ------------------------------------------------------------------------------


XLSX_DIR = get_config("goyderbess", "spec_path")
XLSX_PATH_CSR = os.path.join(XLSX_DIR,"GO_Spec_CSR_300.xlsx")
XLSX_PATH_DMAT = os.path.join(XLSX_DIR,"GO_Spec_DMAT.xlsx")
XLSX_PATH_DMAT_CRG = os.path.join(XLSX_DIR,"GO_Spec_DMAT_CRG.xlsx")
XLSX_PATH_TESTING = os.path.join(XLSX_DIR,"GO_Spec_FLATTEST.xlsx")
PROJECT_DIR = get_config("goyderbess", "project_directory")

REMOVE_INIT_TIME = True


#Run PSCAD DMAT clauses
RUN_STUDIES = True 
if RUN_STUDIES:
    USE_VSLACK_CACHE = True
    CALC_TOV_SHUNT_VAR = True  #SET THIS TO FALSE, TOV values already calculated. 
    volley_size = 16
    plotter = goyderbessPscadPlotter(remove_first_seconds=REMOVE_INIT_TIME)
    MODEL_DIR = os.path.join(PROJECT_DIR, r"""PSCAD\TeslaMP3_PSCAD_v25.20_x64""")
    TEMP_DIR = os.path.join(r"""C:\Grid\results\Goyderbess""", "_temp", NOW_STR)
    STUDY_RESULTS_DIR = os.path.join(r"""C:\Grid\results\Goyderbess""", f"_pscad_results_{NOW_STR}")   
    PROJECT_NAME = "TeslaBESS_BESSonly_4hr"
    
    #Filtering spec
    #spec = load_specs_from_xlsx(XLSX_PATH, sheet_names=SHEETS_TO_PROCESS_CSR)
    spec = load_specs_from_multiple_xlsx([XLSX_PATH_CSR,XLSX_PATH_DMAT,XLSX_PATH_DMAT_CRG,XLSX_PATH_TESTING], sheet_names=[SHEETS_TO_PROCESS_CSR,SHEETS_TO_PROCESS_DMAT,SHEETS_TO_PROCESS_DMAT_CRG,SHEETS_TO_PROCESS_FLAT])
    spec = spec[spec["PSCAD"] == True]
    # spec = spec[spec["Test No"] > 55]
    spec = spec[spec["Test No"] == 161]
    # spec = spec[spec["Test No"].isin({99, 97, 108, 144, 72, 12, 6, 42})]
    # spec = spec[spec["Subtest No"] == 1]
    # spec =spec.head(1)
    print(spec)
    # spec.to_csv("spec.csv")
    # sys.exit()

#Run Analysis on selected sheets
CREATE_ANALYSIS = False
if CREATE_ANALYSIS:
    analysis_extension = ".psout" #Options are .psout or .out

    if RUN_STUDIES: #if run studies and create analysis is true, set inputs dirs to study results dirs
        DMAT_INPUTS_DIR = STUDY_RESULTS_DIR
        CSR_INPUTS_DIR = STUDY_RESULTS_DIR
    else: #if run studies is false, manually assign DIRs for analysis
        DMAT_INPUTS_DIR = r"""D:\results\heywoodbess\LATEST_FULL_RUN\LATEST_DMAT_PSCAD"""
        CSR_INPUTS_DIR = r"""D:\results\heywoodbess\LATEST_FULL_RUN"""

    ANALYSIS_OUTPUTS_DIR = os.path.join(r"""D:\results\heywoodbess\analysis""",f"_analysis_results_{NOW_STR}")

    DPDF_CHARACTERISTIC_POINTS = [(-3, 285),(-2.5, 285), (-0.015, 0), (0.015, 0), (2.5, -285), (3, -285)]
    VDROOP_CHARACTERISTIC_POINTS = [(-0.1,112.575),(-0.04, 112.575), (-0.04, 112.575), (0.04, -112.575), (0.04, -112.575),(0.1,-112.575)]
    HVRT_THRESHOLDS = [
            {
                "value": 1.15,
                "withstand_sec": 60,
                "colour": "yellow",
            },
            {
                "value": 1.25,
                "withstand_sec": 1,
                "colour": "orange",
            },
        ]
    LVRT_THRESHOLDS = [
            {
                "value": 0.8,
                "withstand_sec": 21,
                "colour": "yellow",
            },
            {
                "value": 0.4,
                "withstand_sec": 21,
                "colour": "red",
            },
        ]

    
#Replot PSCAD DMAT studies 
REPLOT_PSCAD = False
if REPLOT_PSCAD:
    extension_replot = ".psout"  #Options are .psout or .pkl
    replotter = goyderbessPscadReplotter(remove_first_seconds=REMOVE_INIT_TIME) 
    PLOT_INPUTS_DIR = r"""D:\results\heywoodbess\_pscad_results_20250523_0924_34238135"""
    REPLOT_OUT_DIR = os.path.join(r"D:\results\heywoodbess\replots",f"replots_init_{NOW_STR}") #DIR where you want png and pdf outputs of replots to be located

#Generate appendices for chosen clauses in SHEETS_TO_PROCESS
CREATE_APPENDIX = False 
if CREATE_APPENDIX:
    if RUN_STUDIES:
        PLOT_RESULTS_DIRS = [STUDY_RESULTS_DIR] #list of paths, can point to multiple directories that contains the plots.
    else:
        PLOT_RESULTS_DIRS = [r"""D:\results\heywoodbess\_pscad_results_20250428_1010_32453161"""]
    
    OUTPUT_DIR_DMAT = os.path.join(r"D:\results\heywoodbess\appendices\appendix_dmat") # DEFAULT DIR where appendices should be stored if you do not have column in spec called "Report". You should set it as a default 
    OUTPUT_DIR_CSR = os.path.join(r"D:\results\heywoodbess\appendices\appendix_csr")
    SHEETS_TO_PROCESS_COMBINED = [SHEETS_TO_PROCESS_CSR,SHEETS_TO_PROCESS_DMAT,SHEETS_TO_PROCESS_DMAT_CRG]
    XLSX_PATHS = [XLSX_PATH_CSR,XLSX_PATH_DMAT,XLSX_PATH_DMAT_CRG]
    issued_date = "22nd April 2025"
    revision_no = "0"

    if not os.path.exists(OUTPUT_DIR_DMAT):
        os.makedirs(OUTPUT_DIR_DMAT)
    if not os.path.exists(OUTPUT_DIR_CSR):
        os.makedirs(OUTPUT_DIR_CSR)

#Create report tables for chosen clauses 
CREATE_REPORT_TABLES = False 
if CREATE_REPORT_TABLES:
    SHEETS_TO_PROCESS_TABLES = [SHEETS_TO_PROCESS_CSR,SHEETS_TO_PROCESS_DMAT,SHEETS_TO_PROCESS_DMAT_CRG]
    SPEC_PATHS_TABLES = [XLSX_PATH_CSR,XLSX_PATH_DMAT,XLSX_PATH_DMAT_CRG]
    TABLE_DIR = r"""D:\luke\heywoodbess\scripts\002 report tables"""
    COLMAP_PATHS = [os.path.join(TABLE_DIR, "column_mapping_DMAT.csv"),os.path.join(TABLE_DIR, "column_mapping_CSR.csv")]
    OUTPUT_DIRS_TABLE = [r"""D:\results\heywoodbess\report_tables\tableoutputsDMAT""",r"""D:\results\heywoodbess\report_tables\tableoutputsCSR"""]
    for output_dirs in OUTPUT_DIRS_TABLE:
        if not os.path.exists(output_dirs):
            os.makedirs(output_dirs)



#-------------------------------------------------------------------------------- END OF SETTINGS AND PATHS -----------------------------------------------------------------------------------------------------------------------


if __name__ == "__main__":

    if RUN_STUDIES: 
        run_pscad_studies(
            USE_VSLACK_CACHE=USE_VSLACK_CACHE,
            spec=spec,
            project_name=PROJECT_NAME,
            volley_size=volley_size,
            plotter=plotter,
            XLSX_DIR=XLSX_DIR,
            MODEL_DIR=MODEL_DIR,
            TEMP_DIR=TEMP_DIR,
            RESULTS_DIR=STUDY_RESULTS_DIR)



    if CREATE_ANALYSIS:
        run_analysis_pscad(
            x86=x86,
            extension=analysis_extension,
            ANALYSIS_TO_RUN=ANALYSIS_TO_RUN,
            DPDF_CHARACTERISTIC_POINTS=DPDF_CHARACTERISTIC_POINTS,
            VDROOP_CHARACTERISTIC_POINTS=VDROOP_CHARACTERISTIC_POINTS,
            LVRT_THRESHOLDS=LVRT_THRESHOLDS,
            HVRT_THRESHOLDS=HVRT_THRESHOLDS,
            DMAT_INPUTS_DIR=DMAT_INPUTS_DIR,
            CSR_INPUTS_DIR=CSR_INPUTS_DIR,
            OUTPUTS_DIR=ANALYSIS_OUTPUTS_DIR)
        

    if REPLOT_PSCAD:
        replot_pscad(
            extension=extension_replot,
            replotter=replotter,
            PLOT_INPUTS_DIR=PLOT_INPUTS_DIR,
            PLOT_OUT_PATH=REPLOT_OUT_DIR,
            x86=x86)


    if CREATE_APPENDIX:
        create_appendix_goyderbess(
                PLOT_RESULTS_DIRS=PLOT_RESULTS_DIRS,
                OUTPUT_DIR_DMAT=OUTPUT_DIR_DMAT,
                OUTPUT_DIR_CSR=OUTPUT_DIR_CSR,
                XLSX_PATHS=XLSX_PATHS,
                SHEETS_TO_PROCESS=SHEETS_TO_PROCESS_COMBINED,
                NOW_STR=NOW_STR,
                issued_date=issued_date,
                revision_no=revision_no,
                x86=x86,
                appendix_project_name=APPENDIX_PROJECT_NAME)


    if CREATE_REPORT_TABLES:
        create_report_table_CSR_DMAT(
            COLMAP_PATHS=COLMAP_PATHS,
            SPEC_PATHS=SPEC_PATHS_TABLES,
            SHEETS_TO_PROCESS=SHEETS_TO_PROCESS_TABLES,
            OUTPUT_DIRS=OUTPUT_DIRS_TABLE,
            x86=x86)