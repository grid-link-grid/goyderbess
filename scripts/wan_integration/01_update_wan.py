from __future__ import with_statement
import sys
import os
import json
import shutil
import logging
import subprocess

MAIN_DIRECTORY = os.getcwd()
sys_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSPY39'  #or where else you find the psspy.pyc
sys.path.append(sys_path_PSSE)
os_path_PSSE=r' C:\Program Files (x86)\PTI\PSSE34\PSSBIN'  # or where else you find the psse.exe
os.environ['PATH'] += ';' + os.environ['PATH']

import psse34
import psspy
from psspy import _i,_f,_s
from typing import List, Union, Optional
import pandas as pd
from tqdm.auto import tqdm
from contextlib import contextmanager
from datetime import datetime
from WANPlotter import WANPlotter
from pallet.psse.Out import Out

from flatrun_wan import flat_run


import redirect
import cattrs

from gridlink.utils.wan.wan_utils_temp import integrate_models,merge_jsons_at_paths,find_files,simple_chandef_to_plotdef, get_network_state_by_chandefs
from pallet.psse.initialiser.Initialiser import Initialiser
from pallet.psse.case.CaseDefinition import CaseDefinition
from pallet.psse import system, statics
from pallet.specs import load_specs_from_xlsx
from pallet.utils.time_utils import now_str


def copy_file(source_file: str, destination_file: str):
    """
    Copies a file from the source path to the destination path.

    :param source_file: Path to the source file
    :param destination_file: Path to the destination file
    """
    try:
        shutil.copy2(source_file, destination_file)
        print(f"File copied from {source_file} to {destination_file}")
    except FileNotFoundError:
        print(f"Error: The source file {source_file} does not exist.")
    except PermissionError:
        print(f"Error: Permission denied. Could not copy the file.")
    except Exception as e:
        print(f"An error occurred: {e}")



def combine_files(
        input_directory:str,
        output_file: str,
        file_type: str,
        remove_originals: bool
        ):
    """
    Combines all text files in the specified directory into a single text file.

    :param input_directory: Directory containing the text files to be combined
    :param output_file: Path to the output file
    """
    files_in_dir =  os.listdir(input_directory)
    with open(output_file, 'w') as outfile:
        for filename in files_in_dir:
            if filename.endswith(file_type):
                file_path = os.path.join(input_directory, filename)
                with open(file_path, 'r') as infile:
                    outfile.write(infile.read())
                    outfile.write("\n")  # Add a newline character to separate files
                # if remove_originals:
                #     os.remove(file_path)
    for file_path in [os.path.join(input_directory,filename) for filename in files_in_dir]:
        os.remove(file_path)

    print(f"All {file_type} files combined into {output_file}")
    
def merge_dictionaries(dict1, dict2):
    merged_dict = {}

    # Add all key-value pairs from the first dictionary to the merged dictionary
    for key, value in dict1.items():
        merged_dict[key] = value

    # Add all key-value pairs from the second dictionary
    for key, value in dict2.items():
        if key in merged_dict:
            # If the key exists in the merged dictionary, combine values into a list
            if isinstance(merged_dict[key], list):
                if isinstance(value,list):
                    [merged_dict[key].append(item) for item in value]
                else:
                    merged_dict[key].append(value)
            else:
                if isinstance(value,list):
                    merged_dict[key] = [merged_dict[key]]
                    [[merged_dict[key].append(item) for item in value]]
                else:
                    merged_dict[key] = [merged_dict[key], value]
        else:
            # If the key doesn't exist, add it to the merged dictionary
            merged_dict[key] = value

    return merged_dict

def chandef_to_plotdef(
        chandef_path:str,
        plotdef_subdir:str
        ):
    
    with open(chandef_path, 'r') as json_file:
        chandef = json.load(json_file)

    plotdef = []

    bus_v = {
            "gen":{},
            "wan":{},
            }
    # Sort bus voltage channels
    if "bus_voltage_channels" in chandef:
        if len(chandef["bus_voltage_channels"])!=0:
            for channel in chandef["bus_voltage_channels"]:
                if any(substring in channel["vol_name"] for substring in ["220kV_V_PU"]):
                    if "columns" in bus_v["wan"]:
                        bus_v["wan"]["columns"].append({
                            "channel": channel["vol_name"].upper(),
                            "alias": channel["vol_name"].upper()
                        })
                    else:
                        bus_v["wan"]["columns"]=[{
                            "channel": channel["vol_name"].upper(),
                            "alias": channel["vol_name"].upper()
                        }]

                if any(substring in channel["vol_name"] for substring in ["POC_V_PU"]):
                    if "columns" in bus_v["gen"]:
                        bus_v["gen"]["columns"].append({
                            "channel": channel["vol_name"].upper(),
                            "alias": channel["vol_name"].upper()
                        })
                    else:
                        bus_v["gen"]["columns"]=[{
                            "channel": channel["vol_name"].upper(),
                            "alias": channel["vol_name"].upper()
                        }]
    if "columns" in bus_v["gen"]:
        bus_v["gen"]["title"] = "POC Voltages (pu)"
        bus_v["gen"]["scaling"] = [1.0] * len(bus_v["gen"]["columns"])
        bus_v["gen"]["y_axis_min_range"] = 0.01

    if "columns" in bus_v["wan"]:
        bus_v["wan"]["title"] = "Network Bus Voltages (pu)"
        bus_v["wan"]["scaling"] = [1.0] * len(bus_v["wan"]["columns"])
        bus_v["gen"]["y_axis_min_range"] = 0.01

    #print(bus_v)
    branch_p = {
            "gen":{},
            "wan":{},
            }
    if "branch_pq_channels" in chandef:
        if len(chandef["branch_pq_channels"])!=0:
            for channel in chandef["branch_pq_channels"]:

                if any(substring in channel["p_name"] for substring in ["POC_P_MW"]):
                    if "columns" in branch_p["gen"]:
                        branch_p["gen"]["columns"].append(
                        {
                            "channel": channel["p_name"].upper(),
                            "alias": channel["p_name"].upper()
                        })
                    else:
                        branch_p["gen"]["columns"]=[
                        {
                            "channel": channel["p_name"].upper(),
                            "alias": channel["p_name"].upper()
                        }]
                else:
                    if "columns" in branch_p["wan"]:
                        branch_p["wan"]["columns"].append(
                        {
                            "channel": channel["p_name"].upper(),
                            "alias": channel["p_name"].upper()
                        })
                    else:
                        branch_p["wan"]["columns"]=[
                        {
                            "channel": channel["p_name"].upper(),
                            "alias": channel["p_name"].upper()
                        }]
    if "columns" in branch_p["gen"]:
        branch_p["gen"]["title"] = "POC Active Power (MW)"
        branch_p["gen"]["scaling"] = [1.0] * len(branch_p["gen"]["columns"])
        branch_p["gen"]["y_axis_min_range"] = 1.0

    if "columns" in branch_p["wan"]:
        branch_p["wan"]["title"] = "Network Active Power Flow (MW)"
        branch_p["wan"]["scaling"] = [1.0] * len(branch_p["wan"]["columns"])
        branch_p["wan"]["y_axis_min_range"] = 1.0

    # Sort branch Q channels
    branch_q = {
            "gen":{},
            "wan":{},
            }
    
    
    if "branch_pq_channels" in chandef:
        if len(chandef["branch_pq_channels"])!=0:
            for channel in chandef["branch_pq_channels"]:

                if any(substring in channel["q_name"] for substring in ["POC_Q_MVAR"]):
                    if "columns" in branch_q["gen"]:
                        branch_q["gen"]["columns"].append(
                        {
                            "channel": channel["q_name"].upper(),
                            "alias": channel["q_name"].upper()
                        })
                    else:
                        branch_q["gen"]["columns"]=[
                        {
                            "channel": channel["q_name"].upper(),
                            "alias": channel["q_name"].upper()
                        }]
                else:
                    if "columns" in branch_q["wan"]:
                        branch_q["wan"]["columns"].append(
                        {
                            "channel": channel["q_name"].upper(),
                            "alias": channel["q_name"].upper()
                        })
                    else:
                        branch_q["wan"]["columns"]=[
                        {
                            "channel": channel["q_name"].upper(),
                            "alias": channel["q_name"].upper()
                        }]

    if "columns" in branch_q["gen"]:
        branch_q["gen"]["title"] = "POC Reactive Power (MVAr)"
        branch_q["gen"]["scaling"] = [1.0] * len(branch_q["gen"]["columns"])
        branch_q["gen"]["y_axis_min_range"] = 1.0

    if "columns" in branch_q["wan"]:
        branch_q["wan"]["title"] = "Network Reactive Power Flow (MVAr)"
        branch_q["wan"]["scaling"] = [1.0] * len(branch_q["wan"]["columns"])
        branch_q["wan"]["y_axis_min_range"] = 1.0


    merged_defs = merge_dictionaries(bus_v,branch_p)
    merged_defs = merge_dictionaries(merged_defs,branch_q)

    #print(merged_defs)

    for key,plotdef in merged_defs.items():
        #print(plotdef)
        plotdef = [item for item in plotdef if item]
        #print(plotdef)
        path = os.path.join(
                            plotdef_subdir,
                            f"{key}.plotdef"
                            )
        with open(path, 'w') as file:
            #print(plotdef)
            json.dump(plotdef, file, indent=4)    


def plot_outputs(plotdef_sub_dir: str,
         outfile_path: str,
         results_sub_dir: str):

    plot_def_paths = find_files(plotdef_sub_dir,".plotdef")
    #print(plot_def_paths)

    filename = os.path.basename(outfile_path).replace(".out","")
    df_from_out = Out(outfile_path).to_df()

    for i,path in enumerate(plot_def_paths):

        plotdef_name = os.path.basename(path).replace(".plotdef","")
        with open(path, 'r') as json_file:
            plot_definitions = json.load(json_file)
        plotter = WANPlotter(
                plot_definitions=plot_definitions,
                remove_first_seconds=0.0,
                )
        plotter.plot_from_df_and_dict(
            df = df_from_out,
            png_path=os.path.join(results_sub_dir,f"{filename}_{plotdef_name}.png"),
            pdf_path=os.path.join(results_sub_dir,f"{filename}_{plotdef_name}.pdf")
        )

def apply_modifications(case_path: str,
                        mods_sub_dir: str):

    psspy.psseinit(20000)
    psspy.case(case_path)

    mods = find_files(mods_sub_dir,".py")
    #assumes order of run is not important
    for mod in mods:
        with open(mod) as f:
            code = f.read()
            exec(code)

    psspy.save(case_path)




if __name__ == "__main__":

    START = 30
    SAMPLE_SIZE = 30
    RUN_TIME_SECS = 2.0

    slack_bus_num = 999
    dyr_min_bus_num = 50

    convergence_tolerance = 100.0
    apply_modifications_flag = True
    skip_models_by_index = []#[10]#[43] # Reduce loyang
    #10, 21,29,43

    gen_def_dir = r"""C:\Users\LukeHyett\OneDrive - Grid-Link\Heywood-PSSE-Wide\WAN\Integration Tool\Generator Definitions"""
    wan_case_dir = r"""C:\Users\LukeHyett\OneDrive - Grid-Link\Heywood-PSSE-Wide\WAN\Snapshots\SpringLo-20241026-110045-SystemNormal\20241026-110045-subTrans-systemNormal_PEC3.sav"""
    # wan_case_dir = r"""C:\Users\LukeHyett\OneDrive - Grid-Link\Learmonth-PSSE-Wide\WAN\Snapshots\AutumnLo_PEC\Model\20240330-120058-subTrans-systemNormal_PEC.sav"""
    snapshot_dll_dir = r""".\dsusr.dll"""
    # snapshot_dll_dir = r"""C:\Users\LukeHyett\Desktop\test\dsusr.dll"""
    snapshot_dyr_dir = r"""C:\Users\LukeHyett\OneDrive - Grid-Link\Heywood-PSSE-Wide\WAN\Snapshots\SpringLo-20241026-110045-SystemNormal\20241026-110045-subTrans-systemNormal_PEC.dyr"""
    # snapshot_dyr_dir = r"""C:\Users\LukeHyett\OneDrive - Grid-Link\Learmonth-PSSE-Wide\WAN\Snapshots\AutumnLo_PEC\Model\DYRs\20240330-120058-subTrans-systemNormal.dyr"""
    results_dir = r"""C:\Users\LukeHyett\OneDrive - Grid-Link\Heywood-PSSE-Wide\Integration Results"""
    models_dir = r"""C:\Users\LukeHyett\OneDrive - Grid-Link\Heywood-PSSE-Wide\WAN\Integration Tool\Models"""
    integration_order_json_dir = r""".\\integration_order.json"""
    modifications_dir = r""".\\Mods"""
    SHEETS_TO_PROCESS = ["init"]
    XLSX_PATH = os.path.join(os.getcwd(),"init.xlsx")

    CHANDEF_PATH = r""".\Custom DEFs\wam_model.chandef"""



    #smib_mapping_df = pd.read_csv(smib_mapping_path)
    model_sav_dirs = find_files(models_dir,".sav")
    model_dyr_dirs = find_files(models_dir,".dyr")
    model_def_dirs = find_files(gen_def_dir,".json")


    # Since bus numbers are assigned based on the order of paths in folders, it may be useful to instead prescribe an order. 
    # This way bus numbers can remain constant.

    if os.path.exists(integration_order_json_dir):
        with open(integration_order_json_dir, 'r') as json_file:
            integration_order = json.load(json_file)        
        sorted_sav_dirs = sorted(model_sav_dirs, key=lambda x: integration_order.get(x, float('inf')))
        #Update file with assumed order of any models that were not previously defined
        my_dict={}
        for j,path in enumerate(sorted_sav_dirs):
            my_dict[path]=j
        with open(integration_order_json_dir, 'w') as file:
            json.dump(my_dict, file, indent=4) 
    else:
        sorted_sav_dirs = model_sav_dirs
        my_dict = {}
        for j,path in enumerate(model_sav_dirs):
            my_dict[path]=j
        with open(integration_order_json_dir, 'w') as file:
            json.dump(my_dict, file, indent=4) 

    debug_dir = os.path.join(results_dir,now_str()+"_debug_data")
    if not os.path.exists(debug_dir):
        os.makedirs(debug_dir)
    with open(os.path.join(debug_dir,"flatrun_log.txt"), "w") as flatrun_log:
            

        for i in list(range(START,SAMPLE_SIZE+1)):
            iteration = list(range(START,SAMPLE_SIZE+1)).index(i)+1
           
            TIME = datetime.now().strftime('%Y-%m-%d_%H-%M-%S-%f')[:-3]
            integrated_case_dir = os.path.join(results_dir,TIME + "_Integration"+"_"+str(min(len(sorted_sav_dirs)-1,i-1)))
            
            sav_dirs = sorted_sav_dirs[:min(len(sorted_sav_dirs),i)]# filtered_model_sav_dirs[:i-len(skip)]
            model_added = sav_dirs[-1]
            print(f"{model_added=}")

            # if skip_list is not None:
            #     filtered_sav_dirs = [item for item in sav_dirs if sav_dirs.index(item) not in skip_list]
            #     print(f"Skipping {len(sav_dirs)-len(filtered_sav_dirs)} cases.")
            #     skipped_cases = [item for item in sav_dirs if item not in filtered_sav_dirs]
            #     for sav_case in skipped_cases:
            #         print(f"Skipped: {sav_case}")

  
            integrate_models(
                sav_dirs,
                model_dyr_dirs,
                model_def_dirs,
                wan_case_dir,
                integrated_case_dir,
                slack_bus_num,
                dyr_min_bus_num,
                convergence_tolerance,
                custom_bus_num_start=795642,
                skip_models_at_index=skip_models_by_index,
                preserve_skipped_bus_numbers=True,
                modifications_dir=modifications_dir
            )


            
            ierr = psspy.close_powerflow()
            ierr = psspy.pssehalt_2()
            
            #Copy snapshot DYR into folder
            copy_file(source_file = snapshot_dyr_dir,
                    destination_file = os.path.join(integrated_case_dir,"DYRs",os.path.basename(snapshot_dyr_dir))
                    )
            
            #copy dsusr dll into folder
            copy_file(source_file = snapshot_dll_dir,
                    destination_file = os.path.join(integrated_case_dir,"DLLs",os.path.basename(snapshot_dll_dir))
                    )

            combine_files(
                input_directory=os.path.join(integrated_case_dir,"DYRs"),
                output_file=os.path.join(integrated_case_dir,"DYRs","combined_dyr.dyr"),
                file_type=".dyr",
                remove_originals=True
            )

            #Organise folders for pallet ingest

            if not os.path.exists(os.path.join(integrated_case_dir,"Model")):
                os.makedirs(os.path.join(integrated_case_dir,"Model"))
            copy_file(source_file = integration_order_json_dir,
                    destination_file = os.path.join(integrated_case_dir,"Model",os.path.basename(integration_order_json_dir))
                    )
            shutil.move(os.path.join(integrated_case_dir,"DYRs"),
                        os.path.join(integrated_case_dir,"Model","DYRs"))
            shutil.move(os.path.join(integrated_case_dir,"DLLs"),
                        os.path.join(integrated_case_dir,"Model","DLLs"))

            #sav_files = find_files(integrated_case_dir,".sav")
            case_dir = os.path.join(integrated_case_dir,"Model",os.path.basename(wan_case_dir))
            shutil.move(
                os.path.join(integrated_case_dir,os.path.basename(wan_case_dir)),
                case_dir
                )

            #Create a plotdef for each generator chandef
            plotdef_subdir = os.path.join(integrated_case_dir,"PLOTDEFs")
            if not os.path.exists(plotdef_subdir):
                os.makedirs(plotdef_subdir)
            for chandef in find_files(os.path.join(integrated_case_dir,"CHANDEFs"),".chandef"):
                simple_chandef_to_plotdef(chandef,plotdef_subdir)


            #Update definition files

            #merge chandefs into single file
            chandef_path = os.path.join(integrated_case_dir,"Model","wan.chandef")
            merge_jsons_at_paths(
                            subdir=os.path.join(integrated_case_dir,"CHANDEFs"),
                            file_type=".chandef",
                            output_path=chandef_path)        


            #merge initdefs into single file
            initdef_path = os.path.join(integrated_case_dir,"Model","wan.initdef")
            merge_jsons_at_paths(
                            subdir=os.path.join(integrated_case_dir,"INITDEFs"),
                            file_type=".initdef",
                            output_path=initdef_path)
            
            #merge savdefs into single file
            savdef_path = os.path.join(integrated_case_dir,"Model","wan.savdef")
            merge_jsons_at_paths(
                            subdir=os.path.join(integrated_case_dir,"SAVDEFs"),
                            file_type=".savdef",
                            output_path=savdef_path)
        


            #Attempt initialisation
            # ATTEMPT_INIT = False

            # if ATTEMPT_INIT:
            #     with open(savdef_path, 'r') as f:
            #         case_definition_raw = json.load(f)
            #     savdef = cattrs.structure(case_definition_raw, CaseDefinition)

            #     with open(initdef_path, 'r') as f:
            #         initialiser_definition_raw = json.load(f)
            #     initialiser = cattrs.structure(initialiser_definition_raw, Initialiser)


            #     spec = load_specs_from_xlsx(xlsx_path = XLSX_PATH, sheet_names=SHEETS_TO_PROCESS)

            #     for _, scenario in spec.iterrows():
            #         system.init(buses=80000)
            #         statics.load_sav(case_dir)
            #         ierr = psspy.solution_parameters_4(realar6=10.0, intgar2=200)
            #         system.check_ierr("SOLUTION_PARAMETERS_4", ierr)

            #         initialiser.run(scenario=scenario, case_definition=savdef, locked_taps=False)
            #         init_case_dir = case_dir.replace(".sav","")+f"_{scenario['Name']}.sav"
            #         statics.save_case(case_dir.replace(".sav","")+f"_{scenario['Name']}.sav")  

            #         #Export generator poc flows
            #         get_network_state_by_chandefs(
            #                 chandefs_sub_dir=os.path.join(integrated_case_dir,"CHANDEFs"),
            #                 case_dir=init_case_dir
            #                 )

            if iteration==1:
                prev_dstate_df = None
                initial_dstate_df = None
            else:
                prev_dstate_df = dstate_df
            if iteration==2:
                initial_dstate_df = dstate_df

            # try:
            # Merge in any additional or custom channel definitions
            # Create list of all chandefs to merge
            chandefs_to_merge = [CHANDEF_PATH]
            [chandefs_to_merge.append(path) for path in find_files(os.path.join(integrated_case_dir, "CHANDEFs"),".chandef")]
            
            # Merge each chandef into a single json file
            for j, file in enumerate(chandefs_to_merge):
                if j==0:
                    # Reading from the JSON file
                    with open(file, 'r') as json_file:
                        staged_chandef = json.load(json_file)
                else:
                    with open(file, 'r') as json_file:
                        data = json.load(json_file)
                    
                    staged_chandef = merge_dictionaries(staged_chandef,data)

            merged_chandef_path = os.path.join(integrated_case_dir, "Model","wan.chandef")    
            with open(merged_chandef_path, 'w') as file:
                json.dump(staged_chandef, file, indent=4) 

            output_dir = os.path.join(results_dir,TIME + "_Dynamics"+"_"+str(i-1))
            dstate_df, ERR = flat_run(
                snapshot_dll_dir = r""".\dsusr.dll""",
                case_dir = os.path.join(integrated_case_dir,"Model",os.path.basename(wan_case_dir)),
                dlls_dir = os.path.join(integrated_case_dir,"Model","DLLs"),
                dyrs_dir = os.path.join(integrated_case_dir,"Model","DYRs"),
                output_dir = output_dir,
                RUN_TIME = RUN_TIME_SECS
            )

            #plot_outputs

            # Create generic plotdef for generators and WAN
            merged_chandef_path = os.path.join(integrated_case_dir,"Model","wan.chandef")
            plotdef_subdir = os.path.join(integrated_case_dir,"PLOTDEFs")
            chandef_to_plotdef(merged_chandef_path,plotdef_subdir)

            #Copy in any custom defs -- add later

            # Plot outputs
            plotdef_subdir = os.path.join(integrated_case_dir,"PLOTDEFs")
            for outfile in find_files(output_dir,".out"):
                plot_outputs(plotdef_subdir, outfile, os.path.dirname(outfile))



            dstate_df.to_csv(os.path.join(debug_dir,f"{i-1}_dstate_errors.csv"), index=False)
            if (prev_dstate_df is not None) and (initial_dstate_df is not None):
                flatrun_log.write(f"\n\nCase {i} out of {SAMPLE_SIZE}")
                flatrun_log.write(f"\nModel added to case: {model_added}")
                #DETERMINE WHETHER THERE ARE ADDITIONAL DSTATES COMPARED TO LAST CASE

                merged = pd.merge(initial_dstate_df, dstate_df, how='outer', on=['STATE', 'BUS-SCT'], indicator=True)
                # Get rows that are unique to df1 based on the combination of 'STATE' and 'BUS-SCT'
                #unique_initial_dstate_df = merged[merged['_merge'] == 'left_only'].drop('_merge', axis=1)
                # Get rows that are unique to df2 based on the combination of 'STATE' and 'BUS-SCT'
                diff_init_dstate_df = merged[merged['_merge'] == 'right_only'].drop('_merge', axis=1)

                print("#----------DSTATES introduced by this case and preceeding cases relative to the initial case:")
                flatrun_log.write("\nDSTATES introduced by this case and preceeding cases relative to the initial case:")
                if diff_init_dstate_df.empty:
                    print("\nNONE")
                    flatrun_log.write("NONE")
                else:
                    print(diff_init_dstate_df)
                    diff_init_dstate_df.to_csv(os.path.join(debug_dir,f"{i-1}_diff_init.csv"), index=False)
                    flatrun_log.write(f"\n***{diff_init_dstate_df.shape[0]} errors identified.")
                    flatrun_log.write(f"\nSee csv log {i-1}_diff_init.csv")

                #DETERMINE WHETHER THERE ARE ADDITIONAL DSTATES COMPARED TO LAST CASE

                merged = pd.merge(prev_dstate_df, dstate_df, how='outer', on=['STATE', 'BUS-SCT'], indicator=True)
                # Get rows that are unique to df1 based on the combination of 'STATE' and 'BUS-SCT'
                #unique_prev_dstate_df = merged[merged['_merge'] == 'left_only'].drop('_merge', axis=1)
                # Get rows that are unique to df2 based on the combination of 'STATE' and 'BUS-SCT'
                unique_dstate_df = merged[merged['_merge'] == 'right_only'].drop('_merge', axis=1)

                print("#----------DSTATES introduced by this case:")
                flatrun_log.write("\nDSTATES introduced by this case:")
                if unique_dstate_df.empty:
                    print("NONE")
                    flatrun_log.write("\nNONE")
                else:
                    print(unique_dstate_df)
                    unique_dstate_df.to_csv(os.path.join(debug_dir,f"{i-1}_diff_prev.csv"), index=False)
                    flatrun_log.write(f"\n***{unique_dstate_df.shape[0]} errors identified.")
                    flatrun_log.write(f"\nSee csv log {i-1}_diff_prev.csv")

            # except Exception as e:
            #     print(f"Exception: {e}")
            #     ERR = 1

            # if ERR == 0:
            #     complete = True
            # elif ERR == 1:
            #     #skipped.append(model_added)
            #     print(f"Failed on model in rank {i-1} (refer to integration order).")
            #     print(f"Model:{model_added}")
            #     flatrun_log.write(f"\n*Failed on the {i-1}th model.")
            #     flatrun_log.write(f"\n*Model:{model_added}")
            #     # print(f"{skipped}")
            #     break



 