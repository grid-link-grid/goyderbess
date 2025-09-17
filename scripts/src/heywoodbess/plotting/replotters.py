import sys
import os
import json
#os.environ["POWERKIT_LICENSE_KEY"] = "396119-1D6701-6F484D-30E154-E30B6A-V3"
import pandas as pd 
import numpy as np
import math
import glob
from pathlib import Path
from typing import Optional, Union, Dict, List, Callable
import ast

from pallet import load_specs_from_csv, load_specs_from_xlsx
from pallet import now_str



from pathlib import Path
from icecream import ic
import inspect

from pallet.specs import load_specs_from_xlsx

from pallet import now_str


from tqdm.auto import tqdm

def create_paths(
            extension:str,
            png_paths:list,
            pdf_paths:list,
            sim_paths:list,
            out_path):
        
    study_no = len(sim_paths)
    for i in range(study_no):
        study_name = os.path.basename(sim_paths[i])
        png_name = study_name.replace(extension,".png")
        pdf_name = study_name.replace(extension,".pdf")
        png_paths.append(os.path.join(out_path,png_name))
        pdf_paths.append(os.path.join(out_path,pdf_name))
    return png_paths, pdf_paths



def replot_pscad(
                extension:str,
                replotter: object,
                PLOT_INPUTS_DIR,
                PLOT_OUT_PATH,
                x86:bool):

    if not x86:
        from pallet.pscad.Psout import Psout
        from pallet.pscad.PscadPallet import PscadPallet
        from pallet.pscad.launching import launch_pscad
        from pallet.pscad.validation import ValidatorBehaviour
        from heywoodbess.plotting.convert_to_pkl import psout_to_pkl
    
    print("Replotting PSCAD plots ...")
    if extension == ".pkl":
        psout_to_pkl(x86,PLOT_INPUTS_DIR)

    OUTPUTS_DIR = Path(PLOT_OUT_PATH)
    OUTPUTS_DIR.mkdir(parents=True,exist_ok=True)


    folders = [item for item in os.listdir(PLOT_INPUTS_DIR) if os.path.isdir(os.path.join(PLOT_INPUTS_DIR, item))]

    for folder in folders:
        
        #Initialises the paths for each file type
        sim_paths = []
        json_paths = []
        png_paths = []
        pdf_paths = []
        #String for input and output paths for each study 
        input_folder_path = os.path.join(PLOT_INPUTS_DIR, folder)

        #Will create a new folder if it doesn't exist (same name study) in the output_dir\pscad
        mk_out_folder_path = Path(os.path.join(OUTPUTS_DIR, folder))
        mk_out_folder_path.mkdir(parents=True, exist_ok=True)
        

        #Creates a list of all the files for the particular study
        files = [file for file in os.listdir(input_folder_path) if os.path.isfile(os.path.join(input_folder_path, file))]

        for filename in files:
            in_file_path = os.path.join(input_folder_path, filename)
            if filename.endswith(".pkl") or filename.endswith(".psout"):
                sim_paths.append(in_file_path)

            if filename.endswith(".json"):
                json_paths.append(in_file_path)


        # Creates paths for png and pdfs
        print("Creating png and pdf paths...")
        png_paths,pdf_paths = create_paths(extension,png_paths,pdf_paths,sim_paths,mk_out_folder_path)


        
        #Iterates through each path and regenerates plots for each test within study/folder
        name_err_count = 0
        name_mismatches = []
        num_studies = len(sim_paths)



        for idx in range(len(sim_paths)):
            
            json_name = os.path.splitext(os.path.basename(json_paths[idx]))[0]
            study_name = os.path.splitext(os.path.basename(sim_paths[idx]))[0]
            
            if json_name == study_name:

                if extension == ".psout":
                    df = Psout(sim_paths[idx]).to_df()
                if extension == ".pkl":
                    df = pd.read_pickle(sim_paths[idx])
                # df.to_csv("output.csv")
                    
                percentage = float(((idx+1)/num_studies)*100)
                print("\n")
                print(folder)
                print(f"Progress: {idx+1}/{num_studies} => {percentage:.2f}%")
                if name_err_count > 0:
                    print(f"Study name mismatch count (JSON and PKL): {name_err_count}\n")
    
                with open(json_paths[idx]) as f:
                    scenario_dict = json.load(f)
                

                replotter.plot_from_df_and_dict(
                                                df=df,
                                                scenario_dict=scenario_dict,
                                                png_path=png_paths[idx],
                                                pdf_path=pdf_paths[idx])
            else:
                print("\n")
                print("ERROR: JSON and STUDY NAMES DO NOT MATCH")
                print(folder)
                print(f"Progress: {idx}/{num_studies}")
                name_err_count = name_err_count +1
                print("Study name mismatch count (JSON and PKL):", name_err_count)
                name_mismatches.append(json_name)
                print(f"ERROR in study name: {name_mismatches}\n")
                
    print("Finished Replotting.")
    print("\n")
    return


def replot_psse(
                extension:str,
                replotter: object,
                PLOT_INPUTS_DIR,
                PLOT_OUT_DIR,
                x86:bool):

    if x86:
        from pallet.psse.Out import Out
        from pallet.psse.PssePallet import PssePallet
        from pallet.psse import system, statics
        from pallet.psse.vslack_calculator import calc_vslack
        from pallet.psse.case.Bus import Bus, BusData
        from pallet.psse.case.Machine import Machine, MachineData
        from heywoodbess.plotting.process_and_calc_psse import pre_process_dataframe
        PSSBIN_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSBIN"""
        PSSPY_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSPY39"""
        sys.path.append(PSSBIN_PATH)
        sys.path.append(PSSPY_PATH)
        import psse34
        import psspy
        from heywoodbess.plotting.convert_to_pkl import out_to_pkl
    
    print("Replotting PSSE plots ...")
    if extension == ".pkl":
        out_to_pkl(x86,PLOT_INPUTS_DIR)

    #Lists all folders in the input directory
    folders = [item for item in os.listdir(PLOT_INPUTS_DIR) if os.path.isdir(os.path.join(PLOT_INPUTS_DIR, item))]

    for folder in folders:

        sim_paths = []
        json_paths = []
        png_paths = []
        pdf_paths = []
        
        #String for input and output paths for each study 
        input_folder_path = os.path.join(PLOT_INPUTS_DIR, folder)
    

        #Will create a folder (same name as study) in the output directory
        mk_out_folder_path = Path(os.path.join(PLOT_OUT_DIR, folder))
        mk_out_folder_path.mkdir(parents=True, exist_ok=True)

        #Creates a list of all the files for the particular study
        files = [file for file in os.listdir(input_folder_path) if os.path.isfile(os.path.join(input_folder_path, file))]

        #Creates a list of paths for different file types (.png, .psout) found in the specific study/folder. 
        for filename in files:
            in_file_path = os.path.join(input_folder_path, filename)
            if filename.endswith(".out") or filename.endswith('.pkl'):
                sim_paths.append(in_file_path)
            if filename.endswith(".json"):
                json_paths.append(in_file_path)

        print("Creating png and pdf paths...")
        png_paths,pdf_paths = create_paths(extension,png_paths,pdf_paths,sim_paths,mk_out_folder_path)   

        name_err_count = 0
        name_mismatches = []
        num_studies = len(sim_paths)
        
        #Iterates through each path and regenerates plots for each test within study 
        for idx in range(len(sim_paths)):

            json_name = os.path.splitext(os.path.basename(json_paths[idx]))[0]
            study_name = os.path.splitext(os.path.basename(sim_paths[idx]))[0]

            if json_name == study_name: 
                if extension == ".out":
                    df = Out(sim_paths[idx]).to_df()
                if extension == ".pkl":
                    df = pd.read_pickle(sim_paths[idx])
                percentage = float(((idx+1)/num_studies)*100)
                print("\n")
                print(folder)
                print(f"Progress: {idx+1}/{num_studies} => {percentage:.2f}%")
                if name_err_count > 0:
                    print(f"Study name mismatch count (JSON and PKL): {name_err_count}\n")
    
                with open(json_paths[idx]) as f:
                    scenario_dict = json.load(f)

                replotter.plot_from_df_and_dict(
                    df=df,
                    scenario_dict=scenario_dict,
                    png_path=png_paths[idx],
                    pdf_path=pdf_paths[idx],
                )
            else:
                print("\n")
                print("ERROR: JSON and PKL NAMES DO NOT MATCH")
                print(folder)
                print(f"Progress: {idx}/{num_studies}")
                name_err_count = name_err_count +1
                print("Study name mismatch count (JSON and PKL):", name_err_count)
                name_mismatches.append(json_name)
                print(f"ERROR in study name: {name_mismatches}\n")

    print("Finished Replotting.")
    print("\n")
    return

def overlay_wan(
                extension:str,
                replotter: object,
                PLOT_INPUTS_DIR_IS,
                PLOT_INPUTS_DIR_OOS,
                PLOT_OUT_DIR,
                x86:bool):

    if x86:
        from pallet.psse.Out import Out
        from pallet.psse.PssePallet import PssePallet
        from pallet.psse import system, statics
        from pallet.psse.vslack_calculator import calc_vslack
        from pallet.psse.case.Bus import Bus, BusData
        from pallet.psse.case.Machine import Machine, MachineData
        from heywoodbess.plotting.process_and_calc_psse import pre_process_dataframe
        PSSBIN_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSBIN"""
        PSSPY_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSPY39"""
        sys.path.append(PSSBIN_PATH)
        sys.path.append(PSSPY_PATH)
        from heywoodbess.plotting.convert_to_pkl import out_to_pkl
    
    print("Replotting PSSE plots ...")
    if extension == ".pkl":
        out_to_pkl(x86,PLOT_INPUTS_DIR_IS)
        out_to_pkl(x86,PLOT_INPUTS_DIR_OOS)

    #Lists all folders in the input directory
    folders1 = [item for item in os.listdir(PLOT_INPUTS_DIR_IS) if os.path.isdir(os.path.join(PLOT_INPUTS_DIR_IS, item))]
    folders2 = [item for item in os.listdir(PLOT_INPUTS_DIR_OOS) if os.path.isdir(os.path.join(PLOT_INPUTS_DIR_OOS, item))]
    folders = list(set(folders1) & set(folders2))

    for folder in folders:

        sim_paths_IS = []
        sim_paths_OOS = []
        json_paths_IS = []
        json_paths_OOS = []
        png_paths = []
        pdf_paths = []
        
        #String for input and output paths for each study 
        input_folder_path_IS = os.path.join(PLOT_INPUTS_DIR_IS, folder)
        input_folder_path_OOS = os.path.join(PLOT_INPUTS_DIR_OOS, folder)

        #Will create a folder (same name as study) in the output directory
        mk_out_folder_path = Path(os.path.join(PLOT_OUT_DIR, folder))
        mk_out_folder_path.mkdir(parents=True, exist_ok=True)

        #Creates a list of all the files for the particular study
        files_IS = [file for file in os.listdir(input_folder_path_IS) if os.path.isfile(os.path.join(input_folder_path_IS, file))]
        files_OOS = [file for file in os.listdir(input_folder_path_OOS) if os.path.isfile(os.path.join(input_folder_path_OOS, file))]

        #Creates a list of paths for different file types (.png, .psout) found in the specific study/folder. 
        for filename in files_IS:
            in_file_path = os.path.join(input_folder_path_IS, filename)
            if filename.endswith(".out") or filename.endswith('.pkl'):
                sim_paths_IS.append(in_file_path)
            if filename.endswith(".json"):
                json_paths_IS.append(in_file_path)

        for filename in files_OOS:
            in_file_path = os.path.join(input_folder_path_OOS, filename)
            if filename.endswith(".out") or filename.endswith('.pkl'):
                sim_paths_OOS.append(in_file_path)
            if filename.endswith(".json"):
                json_paths_OOS.append(in_file_path)

        print("Creating png and pdf paths...")
        png_paths,pdf_paths = create_paths(extension,png_paths,pdf_paths,sim_paths_IS,mk_out_folder_path)   

        name_err_count = 0
        name_mismatches = []
        num_studies = len(sim_paths_IS)
        
        #Iterates through each path and regenerates plots for each test within study 
        for idx in range(len(sim_paths_IS)):

            json_name_IS = os.path.splitext(os.path.basename(json_paths_IS[idx]))[0]
            study_name_IS = os.path.splitext(os.path.basename(sim_paths_IS[idx]))[0]
            json_name_OOS = os.path.splitext(os.path.basename(json_paths_OOS[idx]))[0]
            study_name_OOS = os.path.splitext(os.path.basename(sim_paths_OOS[idx]))[0]

            if json_name_IS == study_name_IS: 
                if extension == ".out":
                    df_IS = Out(sim_paths_IS[idx]).to_df()
                if extension == ".pkl":
                    df_IS = pd.read_pickle(sim_paths_IS[idx])
                percentage = float(((idx+1)/num_studies)*100)
                print("\n")
                print(folder)
                print(f"Progress: {idx+1}/{num_studies} => {percentage:.2f}%")
                if name_err_count > 0:
                    print(f"Study name mismatch count (JSON and PKL): {name_err_count}\n")
    
                with open(json_paths_IS[idx]) as f:
                    scenario_dict = json.load(f)
            else:
                print("\n")
                print("ERROR: JSON and PKL NAMES DO NOT MATCH")
                print(folder)
                print(f"Progress: {idx}/{num_studies}")
                name_err_count = name_err_count +1
                print("Study name mismatch count (JSON and PKL):", name_err_count)
                name_mismatches.append(json_name_IS)
                print(f"ERROR in study name: {name_mismatches}\n")

            if json_name_OOS == study_name_OOS: 
                if extension == ".out":
                    df_OOS = Out(sim_paths_OOS[idx]).to_df()
                if extension == ".pkl":
                    df_OOS = pd.read_pickle(sim_paths_OOS[idx])
                percentage = float(((idx+1)/num_studies)*100)
                print("\n")
                print(folder)
                print(f"Progress: {idx+1}/{num_studies} => {percentage:.2f}%")
                if name_err_count > 0:
                    print(f"Study name mismatch count (JSON and PKL): {name_err_count}\n")
    
                with open(json_paths_OOS[idx]) as f:
                    scenario_dict = json.load(f)

                replotter.plot_from_df_and_dict(
                    df1=df_IS,
                    df2=df_OOS,
                    scenario_dict=scenario_dict,
                    png_path=png_paths[idx],
                    pdf_path=pdf_paths[idx],
                )
            else:
                print("\n")
                print("ERROR: JSON and PKL NAMES DO NOT MATCH")
                print(folder)
                print(f"Progress: {idx}/{num_studies}")
                name_err_count = name_err_count +1
                print("Study name mismatch count (JSON and PKL):", name_err_count)
                name_mismatches.append(json_name_OOS)
                print(f"ERROR in study name: {name_mismatches}\n")

    print("Finished Overlaying.")
    print("\n")
    return
