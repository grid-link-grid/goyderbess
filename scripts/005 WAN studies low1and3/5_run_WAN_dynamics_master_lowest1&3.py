import sys
import os
from pallet.utils.time_utils import now_str
from pathlib import Path

PSSBIN_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSBIN"""
PSSPY_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSPY39"""
sys.path.append(PSSBIN_PATH)
sys.path.append(PSSPY_PATH)

import os, glob

from gridlink.utils.wan.wan_utils import find_files

from heywoodbess.plotting.heywoodbessWANDynamicsPlotter import heywoodbessWANPlotter
# from heywoodbess.plotting.heywoodbessWANDynamicsPlotterOOS import CGBESSWANPlotterOOS
# from heywoodbess.plotting.heywoodbessWANDynamicsOverlayer import CGBESSWANOverlayer


from heywoodbess.plotting.process_and_calc_wan_dynamics import pre_process_dataframe
from heywoodbess.studyrunners.psse_wan_study_runner import run_psse_studies
from heywoodbess.analysis.run_analysis_psse import run_analysis_psse
# from heywoodbess.utils.vpoc import calc_vpocs


from heywoodbess.plotting.replotters import overlay_wan
from heywoodbess.plotting.replotters import replot_psse
from heywoodbess.appendices.create_appendix import create_appendix_heywoodbess
from heywoodbess.utils.create_report_tables_spec import create_report_table

from pallet.specs import load_specs_from_csv, load_specs_from_xlsx

SHEETS_TO_PROCESS = [
    # "Line Contingencies",
    # "Other Contigencies",
    # "52513_PFref_HighestLoadCase",
    # "52513_Qref_HighestLoadCase",
    # "52513_Vref_HighestLoadCase",
    "52513_PFref_LowestLoadCase1&3",
    "52513_Qref_LowestLoadCase1&3",
    "52513_Vref_LowestLoadCase1&3",
    # "52513_PFref_LowestLoadCase2&4",
    # "52513_Qref_LowestLoadCase2&4",
    # "52513_Vref_LowestLoadCase2&4"
    ]

ANALYSIS_TO_RUN = [
                        

# # 5.13
#             "Vgrid step analysis",              
            "Vref step analysis",              
            "Qref step analysis",                
            "PFref step analysis",
            # "IQ Rise Settle & P Recovery Curve",    
# # # 5.2.5.8
#             "S5258 Active Power Reduction"        
# # 5.5
            #"diq/dV characteristic",                
                     

            ]
 


NOW_STR = now_str()

x86 = True

SAV_VERSION = "v1-1-0"
DYR_VERSION = "v1-1-0"

CALC_VPOCS = False
SKIP_INITIALISATION = True

MODEL_DIR = r"""D:\inputs\heywood\WAN\Lowest_case_1&3\Model"""#r"""C:\Users\JaredGeere\Grid-Link\Projects - Documents\Enzen\CGBESS\02_Deliverables\029 WAN Cases\3_BK_Integration\Finalised cases\Case 2"""

if not SKIP_INITIALISATION:
    INITDEF_PATH = os.path.join(MODEL_DIR,"wan_simplified.initdef")
else:
    INITDEF_PATH = os.path.join(MODEL_DIR,"skip_init.initdef")

FILENAME = "20241026-110045-subTrans-systemNormal_PEC3.sav"

SAV_CASE_PATH = os.path.join(MODEL_DIR,FILENAME)

CHANDEF_PATH = os.path.join(MODEL_DIR,"CHANDEFs")
SAVDEF_PATH = os.path.join(MODEL_DIR,"SAVDEFs")

XLSX_PATH_ENV = r"""C:\Users\abai\Grid-Link\Projects - Atmos\Heywood BESS R0\02_Deliverables\PROJECT_SPEC"""#r"""C:\Users\JaredGeere\Grid-Link\Projects - Documents\Enzen\CGBESS\02_Deliverables\028 R1 Package\05 WAN Spec"""#os.getenv("GESF_SPEC_PATH")
XLSX_PATH = os.path.join(XLSX_PATH_ENV, "HY_Spec_WAN.xlsx")
#----------------------------------------------------------------------------- SET RUN SETTINGS + PATHS  ------------------------------------------------------------------------------
RUN_STUDIES = True
if RUN_STUDIES:

    plotter = heywoodbessWANPlotter(pre_process_fn=pre_process_dataframe)#CGBESSWANPlotterOOS()#CGBESSWANPlotter(pre_process_fn=pre_process_dataframe)
    RESULTS_DIR = os.path.join(r'''D:\results\heywoodbess\wan''',f"{SAV_VERSION}_sav_{DYR_VERSION}_dyr", now_str())
    
    if not os.path.exists(RESULTS_DIR):
        os.makedirs(RESULTS_DIR)

    spec = load_specs_from_xlsx(XLSX_PATH, sheet_names=SHEETS_TO_PROCESS, add_sheet_name_column=True)
    # spec = spec[spec["Test No"]==5]
    # spec = spec[spec["Test No"].isin([8,10,11,12,13,14,15,16])]

RUN_ANALYSIS = False
if RUN_ANALYSIS:
    analysis_extension = ".out"
    DMAT_INPUTS_DIR = r"""C:\Grid\cg\psse\psse_results\dmat"""
    CSR_INPUTS_DIR = r"""D:\jared\temp"""#r"""D:\Latest results for presentation\CG BESS\PSSE WAN\Case 2\IS"""#r"""D:\Latest results for presentation\CG BESS\PSSE WAN\Case 2\IS"""#r"""D:\results\heywoodbess\wan\v1-1-0_sav_v1-1-0_dyr\20250508_1612_34380371"""#r"""C:\Grid\results\heywoodbess\WAN\v1-1-0_sav_v1-1-0_dyr\20250507_1417_18688500"""#r"""C:\Grid\results\heywoodbess\WAN\v1-1-0_sav_v1-1-0_dyr\20250506_2307_16467499"""
    ANALYSIS_OUTPUTS_DIR = os.path.join(r'''D:\results\heywoodbess\wan\psse_analysis''',f"_analysis_{NOW_STR}")

    DPDF_CHARACTERISTIC_POINTS = [(-3, 250), (-2.015, 250), (-0.015, 0), (0.015, 0), (2.015, -250), (3, -250)] #Updated for GESF
    VDROOP_CHARACTERISTIC_POINTS = [(-0.1,98.75),(-0.0395, 98.75), (0.0395, -98.75), (0.1,-98.75)] # Updated for GESF

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
    replotter = heywoodbessWANPlotter(pre_process_fn=pre_process_dataframe)                                        
    PLOT_INPUTS_DIR = r"""C:\Grid\Heywood work folder\WAN\results\WAN STUDY\v1-1-0_sav_v1-1-0_dyr\20250728_0948_34156985"""
    PLOT_OUT_PATH = os.path.join(r"""C:\Grid\Heywood work folder\WAN\results\WAN STUDY""",f"replot_{NOW_STR}") #DIR where you want png and pdf outputs of replots to be located


#Overlay IS and OOS WAN studies 
OVERLAY = False
if OVERLAY:
    extension_replot = ".out"
    replotter = heywoodbessWANPlotter(pre_process_fn=pre_process_dataframe)                                        
    PLOT_INPUTS_DIR_IS = r"""D:\Latest results for presentation\CG BESS\PSSE WAN\1_Simulations\Case 5\IS"""
    PLOT_INPUTS_DIR_OOS = r"""D:\Latest results for presentation\CG BESS\PSSE WAN\1_Simulations\Case 5\OOS"""
    PLOT_OUT_PATH = os.path.join(r"""D:\Latest results for presentation\CG BESS\PSSE WAN\3_Overlays\Case 5""",f"replot_{NOW_STR}") #DIR where you want png and pdf outputs of replots to be located
    SPEC_PATH=XLSX_PATH
    SHEETS_TO_PROCESS=['Case2_Contingencies']

#Generate appendices for chosen clauses 
CREATE_APPENDIX = False
if CREATE_APPENDIX:
    PLOT_RESULTS_DIRS = [os.path.join(r'D:\Latest results for presentation\CG BESS\PSSE WAN\3_Overlays\Case 5\replot_20250514_1533_30987292')] # PLOT_OUT_PATH 
    OUTPUT_DIR_DMAT = os.path.join(r"D:\dan\cg\pscad\appendices\appendix_dmat") # DEFAULT DIR where appendices should be stored if you do not have column in spec called "Report". You should set it as a default 
    OUTPUT_DIR_CSR = os.path.join(r"D:\dan\cg\pscad\appendices\appendix_csr")
    SHEETS_TO_PROCESS_COMBINED = [SHEETS_TO_PROCESS]
    XLSX_PATHS = [XLSX_PATH]
    issued_date = "16 May 2025"
    revision_no = "1-0-0"

    #want to ignore the spec and overwrite title and doc number?
    overwrite_title = True
    title = "Appendix T WAN Case 4 - Wide Area Network Contingency Overlays [S5.2.5.12]"
    report_no = 5 # must be an int

    if not os.path.exists(OUTPUT_DIR_DMAT):
        os.makedirs(OUTPUT_DIR_DMAT)
    if not os.path.exists(OUTPUT_DIR_CSR):
        os.makedirs(OUTPUT_DIR_CSR)
    

#Create report tables for chosen clauses 
CREATE_REPORT_TABLES = False 
if CREATE_REPORT_TABLES:
    SPEC_PATH = XLSX_PATH
    COLMAP_PATH = os.path.join(r"""C:\Grid\heywoodbess\scripts\013_generate_report_tables""", "column_mapping_1.csv")
    OUTPUT_DIR = Path(r"""C:\Grid\cg\psse\table_outputs""")
    #OUTPUT_DIR.mkdir(parents=True,exist_ok=True)


#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
if __name__ == "__main__":

    if RUN_STUDIES:
        run_psse_studies(
            spec=spec,
            MODEL_DIR=MODEL_DIR,
            plotter=plotter,
            RESULTS_DIR=RESULTS_DIR,
            smib_study=False,
            skip_initialisation=SKIP_INITIALISATION,
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

    if OVERLAY:
        overlay_wan(
            extension=extension_replot,
            replotter=CGBESSWANOverlayer(),
            PLOT_INPUTS_DIR_IS=PLOT_INPUTS_DIR_IS,
            PLOT_INPUTS_DIR_OOS=PLOT_INPUTS_DIR_OOS,
            PLOT_OUT_DIR=PLOT_OUT_PATH,
            SPEC_PATH=XLSX_PATH,
            SHEETS_TO_PROCESS=SHEETS_TO_PROCESS,
            x86=x86)


    if CREATE_APPENDIX:
        create_appendix_cgbess(
                PLOT_RESULTS_DIRS=PLOT_RESULTS_DIRS,
                OUTPUT_DIR_DMAT=OUTPUT_DIR_DMAT,
                OUTPUT_DIR_CSR=OUTPUT_DIR_CSR,
                XLSX_PATHS=XLSX_PATHS,
                SHEETS_TO_PROCESS=SHEETS_TO_PROCESS_COMBINED,
                NOW_STR=NOW_STR,
                issued_date=issued_date,
                revision_no=revision_no,
                overwrite_title=overwrite_title,
                hard_title=title,
                report_no=report_no,
                x86=x86)


    if CREATE_REPORT_TABLES:
        create_report_table(
            COLMAP_PATH=COLMAP_PATH,
            SPEC_PATH=SPEC_PATH,
            SHEETS_TO_PROCESS=SHEETS_TO_PROCESS,
            OUTPUT_DIR=OUTPUT_DIR)

    if CALC_VPOCS:

        spec = load_specs_from_xlsx(XLSX_PATH, sheet_names=SHEETS_TO_PROCESS, add_sheet_name_column=True)

        calc_vpocs(
                poc_bus_num=2000,
                case_dir=SAV_CASE_PATH,
                savdef_path=SAVDEF_PATH,
                initdef_path=INITDEF_PATH,
                specification=spec)