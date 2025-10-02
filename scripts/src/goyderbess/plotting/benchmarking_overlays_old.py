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

PSSBIN_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSBIN"""
PSSPY_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSPY39"""
sys.path.append(PSSBIN_PATH)
sys.path.append(PSSPY_PATH)

import psse34
import psspy

from pallet.psse.Out import Out
from pallet import now_str

from gridlink.utils.wan.wan_utils import find_files
from goyderbess.plotting.goyderbessBenchmarkPlotter import goyderbessBenchmarkPlotter

from pallet.psse.Out import Out
from pallet import now_str

from gridlink.utils.wan.wan_utils import find_files



def run_benchmark_overlays( 
            comparison_subdirs:list,
            include_error_bar:list,
            benchmarking_plotter: object,
            PSCAD_INPUT_DIR,
            PSSE_INPUT_DIR,
            RESULTS_DIR):


    #EXTRACT NECESSARY COLUMNS TO IMPROVE SPEED
    for folder in comparison_subdirs:

        if folder != 'Fgrid Steps':
            PSCAD_COLUMNS = ["Ppoc_MW","Qpoc_MVAr","Vrms_poc_pu","PPC_FRT_FLAG","INV_FRT_FLAG"]
            PSSE_COLUMNS = ["PPOC_MW","QPOC_MVAR","VRMS_POC_PU","PPC_FRT_ACTIVE","INV1_FRT_STATE"]
        else:
            PSCAD_COLUMNS = ["Ppoc_MW","Qpoc_MVAr","Vrms_poc_pu","Fslack_meas_Hz"]
            PSSE_COLUMNS = ["PPOC_MW","QPOC_MVAR","VRMS_POC_PU","POC_F_PU"]


        PSCAD_INPUT_SUBDIR = os.path.join(PSCAD_INPUT_DIR,folder)
        PSSE_INPUT_SUBDIR = os.path.join(PSSE_INPUT_DIR,folder)



        PSCAD_extension = ".pkl"
        PSSE_extension = ".out"

        psout_paths = find_files(PSCAD_INPUT_SUBDIR,PSCAD_extension)


        pscad_tests = [os.path.basename(path) for path in psout_paths]

        out_paths = find_files(PSSE_INPUT_SUBDIR,PSSE_extension)
        psse_tests = [os.path.basename(path) for path in out_paths]

        equivalent_psse_paths = [pscad_path.replace(PSCAD_extension,PSSE_extension) for pscad_path in psout_paths]

        # print("psse:",set([test.replace(PSSE_extension,"") for test in psse_tests]))
        # print("pscad:",set([test.replace(PSCAD_extension,"") for test in pscad_tests]))


        common_tests = list(
                            set([test.replace(PSSE_extension,"") for test in psse_tests])
                            & 
                            set([test.replace(PSCAD_extension,"") for test in pscad_tests])
                            )
        #print(common_tests)


        for i in tqdm(range(len(common_tests)), desc=f"PSCAD PSSE Overlays"):
            

            psse_path = [s for s in out_paths if common_tests[i] in s]
            pscad_path = [s for s in psout_paths if common_tests[i] in s]
            

            if len(psse_path) > 1:
                print(f"Error - too many psse matches were made: {psse_path}")
            else:
                psse_path = psse_path[0]

            if len(pscad_path) > 1:
                print(f"Error - too many pscad matches were made: {pscad_path}")
            else:
                pscad_path = pscad_path[0]
                
            dict = pscad_path.replace(PSCAD_extension,".json")    
            with open(dict) as f:
                scenario_dict = json.load(f)

            psout_df = pd.read_pickle(pscad_path)
            # cols_df = pd.DataFrame(psout_df.columns, columns=["pscad signals"])
            # cols_df.to_csv("pscad signals.csv")

            # print(f"These are the pscad signals: {psout_df.columns}")
            psout_df = psout_df[PSCAD_COLUMNS]
            
            # if folder != "Fgrid Steps":
            #     psout_df['INV1_IN_FRT'] = psout_df['PCU1_FRT_STATE'].apply(lambda x: 1 if x == 99 else 0)
             

            #ADD ERROR BANDS
            if folder in include_error_bar:

                psout_df["MW_UPPER"] = psout_df["Ppoc_MW"] + 0.1*psout_df["Ppoc_MW"]
                psout_df["MW_LOWER"] = psout_df["Ppoc_MW"] - 0.1*psout_df["Ppoc_MW"]

                psout_df["kV_UPPER"] = psout_df["Vrms_poc_pu"] + 0.1*psout_df["Vrms_poc_pu"]
                psout_df["kV_LOWER"] = psout_df["Vrms_poc_pu"] - 0.1*psout_df["Vrms_poc_pu"]

                psout_df["MVAr_UPPER"] = psout_df["Qpoc_MVAr"] + 0.1*psout_df["Qpoc_MVAr"]
                psout_df["MVAr_LOWER"] = psout_df["Qpoc_MVAr"] - 0.1*psout_df["Qpoc_MVAr"]

            REMOVE_FIRST_SECONDS = float(scenario_dict["substitutions"]["TIME_Full_Init_Time_sec"])

            psout_df = psout_df[psout_df.index > REMOVE_FIRST_SECONDS]
            psout_df.index = psout_df.index - REMOVE_FIRST_SECONDS
            
            #print(psout_df)

            out_df = Out(psse_path).to_df()
            # cols_df = pd.DataFrame(out_df.columns, columns=["psse Signals"])
            # cols_df.to_csv("psse signals.csv")
            out_df = out_df[PSSE_COLUMNS]
            out_df = out_df[out_df.index >= 0]
            out_df = out_df.rename(columns=lambda x: "PSSE_" + x)

            #print(out_df)


            #print(pscad_path)
            #print(psse_path)
            # len_psout_df = len(psout_df)
            # len_out_df = len(out_df)

            # # Calculate the difference
            # difference = abs(len_psout_df - len_out_df)

            # # Trim rows only from the DataFrame with more rows
            # if len_psout_df < len_out_df:
            #     out_df = out_df.iloc[difference:]  # Remove rows from the start of df B
            # else:
            #     psout_df = psout_df.iloc[difference:]  # Remove rows from the start of df A


            ###------------- SAMPLING -------------------------

            # print(psout_df.index[:5])

            # print(out_df.index[:5])
            
            # # Assume DataFrame A has high-resolution data
            time_index_a = pd.date_range(start="2023-01-01", periods=len(psout_df), freq="1ms")  # Example frequency
            psout_df.index = time_index_a

            # Assume DataFrame B has lower resolution data
            time_index_b = pd.date_range(start="2023-01-01", periods=len(out_df), freq="1ms")  # Example frequency
            out_df.index = time_index_b


            # Resample DataFrame A (high-resolution) to 1-second intervals
            psout_df_resampled = psout_df.resample("1ms").mean()  # Using mean to aggregate

            # Resample DataFrame B (low-resolution) to 1-second intervals (if needed)
            out_df_resampled = out_df.resample("1ms").mean()  # Using interpolate to fill missing data


            psout_df_resampled.index = psout_df_resampled.index.astype('int64') / 1e9  # Convert to seconds
            psout_df_resampled.index = psout_df_resampled.index - psout_df_resampled.index[0]

            out_df_resampled.index = out_df_resampled.index.astype('int64') / 1e9  # Convert to seconds
            out_df_resampled.index = out_df_resampled.index - out_df_resampled.index[0]


            filename = os.path.basename(pscad_path.replace(PSCAD_extension,""))
            sub_folder = os.path.basename(os.path.dirname(pscad_path))


            RESULTS_PATH = os.path.join(RESULTS_DIR,sub_folder)

            if not os.path.exists(RESULTS_PATH):
                os.makedirs(RESULTS_PATH)

            

            png_path = os.path.join(RESULTS_PATH,f"{filename}.png")
            pdf_path = os.path.join(RESULTS_PATH,f"{filename}.pdf")

            print(f"File created: {png_path}")
            benchmarking_plotter.plot_benchmarking(
                study_name = folder,
                error_studies = include_error_bar,
                df_psse = out_df_resampled,
                df_pscad = psout_df_resampled,
                scenario_dict=dict,
                png_path=png_path,
                pdf_path=pdf_path
            )

    return