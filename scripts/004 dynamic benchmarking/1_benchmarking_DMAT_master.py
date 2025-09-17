import os
import glob
import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import json
import sys
from tqdm.auto import tqdm
from pathlib import Path

from pallet import now_str

from heywoodbess.plotting.heywoodbessBenchmarkPlotter import heywoodbessBenchmarkPlotter
from heywoodbess.plotting.benchmarking_overlays import run_benchmark_overlays
from heywoodbess.appendices.create_benchmark_appendix import create_appendix_benchmarking
from heywoodbess.utils.create_report_tables_spec import create_report_table_benchmarking
from heywoodbess.utils.convert_to_pkl import convert_to_pkl


NOW_STR = now_str()

SHEETS_TO_PROCESS = [


#--------------------- DMAT ------------------------#
    
    # "3211_Pref", 
    # "3210_Qref",
    # "3210_PFref",

    # "3212_Fgrid",
    # "324_325_Faults",
    
    # "3210_Vref",

    "3210_3214_Vgrid",
    # "329_TOV",
    # "3216_PhaseSteps",

        ]


#------------------------------------------------------ SET RUN SETTINGS + PATHS --------------------------------------------------------------------------------
MODEL_VERSION = "v0-0-0"
XLSX_PATH = os.path.join(r"C:\Users\lhyett.GRID-LINK\Grid-Link\Projects - PROJECT_SPEC (1)", "HY_Spec_DMAT.xlsx")

PSCAD_INPUT_DIR = r"""D:\results\heywoodbess\_pscad_results_20250625_1705_25632808"""
PSSE_INPUT_DIR = r"""D:\results\heywoodbess\psse\v0-0-4_sav_v0-0-4_dyr\20250625_1642_30536749"""

CONVERT_TO_PKL = False
if CONVERT_TO_PKL:

    CONVERT_PSSE_TO_PKL = True #Set to true if you want .out to become .pkl. Otherwise set to FALSE, to have .psout to pkl
    


#Create benchmarking overlays - PSSE & PSCAD
CREATE_OVERLAYS = True
if CREATE_OVERLAYS:
    ACTIVATE_VENV32 = False

    OUTPUT_DIR = os.path.join(r"""D:\results\heywoodbess\benchmarking""",f"DMAT_BENCHMARKING_{NOW_STR}")
    comparison_subdirs = ['Balanced Faults', 'Fgrid Steps', 'PFref Steps', 'Phase Step', 'Pref Steps', 'Qref Steps', 'TOV', 'Vgrid Steps', 'Vref Steps']
    # comparison_subdirs = ['Vgrid Steps']
    comparison_subdirs = ['Vref Steps', 'PFref Steps', 'Qref Steps']
    error_bar_clauses = ['Vref Steps','Vgrid Steps','Qref Steps','PFref Steps','Pref Steps'] #Clauses that should have error bars

   
    benchmarking_plotter = heywoodbessBenchmarkPlotter()
    
#Create Benchmarking Appendix - PSSE & PSCAD
CREATE_BENCHMARK_APPENIDX = False
if CREATE_BENCHMARK_APPENIDX:
    PLOT_RESULTS_DIR = r"C:\Grid\cg\benchmarking\overlays\DMAT_BENCHMARKING_20250311_0950_23308531" #Path to png plots 
    APPENDIX_OUTPUT_DIR = os.path.join(r"C:\Grid\cg\benchmarking\appendix",f"apdx_{NOW_STR}")
    SEPARATE_CRG_DISCRG_APPENDIX = True

CREATE_REPORT_TABLES = False
if CREATE_REPORT_TABLES:
    SPEC_PATH = XLSX_PATH
    COLMAP_PATH = os.path.join(r"""C:\Grid\cgbess\scripts\013_generate_report_tables""", "column_mapping_bench.csv")
    OUTPUT_DIR = Path(r"""C:\Grid\cg\benchmarking\tables""")
    OUTPUT_DIR.mkdir(parents=True,exist_ok=True)

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------
if __name__ == "__main__":

    if CONVERT_TO_PKL:
        convert_to_pkl(
            convert_psse_to_pkl=CONVERT_PSSE_TO_PKL,
            PSSE_INPUT_DIR=PSSE_INPUT_DIR,
            PSCAD_INPUT_DIR=PSCAD_INPUT_DIR)
        
    if CREATE_OVERLAYS:
        run_benchmark_overlays(
            comparison_subdirs=comparison_subdirs,
            include_error_bar=error_bar_clauses,
            benchmarking_plotter=benchmarking_plotter,
            PSCAD_INPUT_DIR=PSCAD_INPUT_DIR,
            PSSE_INPUT_DIR=PSSE_INPUT_DIR,
            RESULTS_DIR=OUTPUT_DIR,
            x86=ACTIVATE_VENV32
        )
    
    if CREATE_BENCHMARK_APPENIDX:
        create_appendix_benchmarking(
            PLOT_RESULTS_DIR=PLOT_RESULTS_DIR,
            APPENDIX_OUTPUT_DIR=APPENDIX_OUTPUT_DIR,
            SEPARATE_CRG_DISCRG_APPENDIX=SEPARATE_CRG_DISCRG_APPENDIX
        )

    if CREATE_REPORT_TABLES:
        create_report_table_benchmarking(
            COLMAP_PATH=COLMAP_PATH,
            SPEC_PATH=SPEC_PATH,
            SHEETS_TO_PROCESS=SHEETS_TO_PROCESS,
            OUTPUT_DIR=OUTPUT_DIR)