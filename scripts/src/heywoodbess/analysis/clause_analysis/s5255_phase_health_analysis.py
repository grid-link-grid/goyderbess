import os
import json
import pandas as pd
import numpy as np
from typing import List, Callable, Optional
from tqdm.auto import tqdm
from pallet.pscad.Psout import Psout
from pallet.dsl.ParsedDslSignal import ParsedDslSignal
import matplotlib.pyplot as plt


def conduct_phase_voltage_analysis(
        psout_paths : List[str],
        output_png_path: str,
        output_csv_path : str,

        timing_cmd_spec_key : str,
        fault_type_cmd_spec_key: str,
        fault_type_enums: dict,
        healthy_phase_enums: dict,

        phA_channel_key: str,
        phB_channel_key: str,
        phC_channel_key: str,
        groupby_str_for_plot_title: str=None,
        ):
    


    results_df = pd.DataFrame(columns=[
                "Name",
                "Fault Type",
                "In-Fault Voltage Phase A (pu)",
                "In-Fault Voltage Phase B (pu)",
                "In-Fault Voltage Phase C (pu)",
                "In-Fault Max Phase Voltage (pu)",
                "In-Fault Min Phase Voltage (pu)",
                "Post-Fault Voltage Phase A (pu)",
                "Post-Fault Voltage Phase B (pu)",
                "Post-Fault Voltage Phase C (pu)",
                "Post-Fault Max Phase Voltage (pu)",
                "Post-Fault Min Phase Voltage (pu)",
                "Fault Level",
            ])
    

    for i in tqdm(range(len(psout_paths)), desc="Phase Voltage Health Analysis"):
        psout_path = psout_paths[i]
        json_path = psout_path.split(".psout")[0] + ".json" 

        with open(json_path, 'r') as f:
            spec = json.load(f)

        png_output_dir = os.path.dirname(output_png_path)
        csv_output_dir = os.path.dirname(output_csv_path)

        for dir in [png_output_dir,csv_output_dir]:
            if not os.path.exists(dir):
                os.makedirs(dir)

        TIME_SHIFT_SEC = 0.01

        init_time_sec = float(spec["substitutions"]["TIME_Full_Init_Time_sec"])
        filename = spec["File_Name"]
        fault_level = spec["Grid_FL_MVA_sig"]
        df = Psout(psout_path).to_df(secs_to_remove=init_time_sec)


        timing_cmd = ParsedDslSignal.parse(spec[timing_cmd_spec_key]).signal
        fault_type_cmd = ParsedDslSignal.parse(str(spec[fault_type_cmd_spec_key])).signal
        prefault_times = [piece.time_start - TIME_SHIFT_SEC for piece in timing_cmd.pieces if piece.target==1.0]
        in_fault_settled_times = [piece.time_start - TIME_SHIFT_SEC for piece in timing_cmd.pieces if piece.target==0.0]
        

        #If the fault_type_cmd is a varying signal and not a constant
        if len(fault_type_cmd.pieces)!=0:
            list_of_fault_types = [piece.target for piece in fault_type_cmd.pieces if piece.target!=0]
        else:
            list_of_fault_types = [fault_type_cmd.initial_value]

        #If there is just a single fault type, create a list with the same length as the settled_times list
        #This is because we want the two lists to match up with one fault type per fault event time for indexing purposes
        if len(list_of_fault_types)==1:
            list_of_fault_types = list_of_fault_types*len(in_fault_settled_times)
        
    
        for index, fault in enumerate(list_of_fault_types):
            
            #If there is more than one fault event in the psout file, create a suffix for the file name
            if len(list_of_fault_types)>1:
                fault_instance = "_FAULT"+str(index+1)
            else:
                fault_instance = ""

            fault_type_int = list_of_fault_types[index]
            fault_type_str = fault_type_enums[fault_type_int] if fault_type_int in fault_type_enums else "undefined"

            
            snip_df = df[df.index < in_fault_settled_times[index]]

            in_fault_phase_voltages = {
                "Phase A": snip_df[phA_channel_key].iloc[-1],
                "Phase B": snip_df[phB_channel_key].iloc[-1],
                "Phase C": snip_df[phC_channel_key].iloc[-1],
            }

            
            max_in_fault_pu = max([in_fault_phase_voltages[x] for x in healthy_phase_enums[fault_type_int]]) if len(healthy_phase_enums[fault_type_int]) !=0 else 9999
            min_in_fault_pu = min([in_fault_phase_voltages[x] for x in healthy_phase_enums[fault_type_int]]) if len(healthy_phase_enums[fault_type_int]) !=0 else 9999

            post_fault_settled_time = prefault_times[index+1] if index!=len(in_fault_settled_times)-1 else df.index[-1]
            #For post fault we are interested in the max value from fault end to before next fault or to the end of the df 
            snip_df = df[(df.index < post_fault_settled_time) & (df.index > (in_fault_settled_times[index])+TIME_SHIFT_SEC)]

            post_fault_phase_voltages = {
                "Phase A": max(snip_df[phA_channel_key]),
                "Phase B": max(snip_df[phB_channel_key]),
                "Phase C": max(snip_df[phC_channel_key]),
            }

            max_post_fault_pu = max([post_fault_phase_voltages[x] for x in post_fault_phase_voltages])
            min_post_fault_pu = min([post_fault_phase_voltages[x] for x in post_fault_phase_voltages])


            new_row = {
                    "Name": filename+fault_instance,
                    "Fault Type": fault_type_str,
                    "In-Fault Voltage Phase A (pu)": in_fault_phase_voltages["Phase A"],
                    "In-Fault Voltage Phase B (pu)": in_fault_phase_voltages["Phase B"],
                    "In-Fault Voltage Phase C (pu)": in_fault_phase_voltages["Phase C"],
                    "In-Fault Max Phase Voltage (pu)": max_in_fault_pu,
                    "In-Fault Min Phase Voltage (pu)": min_in_fault_pu,
                    "Post-Fault Voltage Phase A (pu)": post_fault_phase_voltages["Phase A"],
                    "Post-Fault Voltage Phase B (pu)": post_fault_phase_voltages["Phase B"],
                    "Post-Fault Voltage Phase C (pu)": post_fault_phase_voltages["Phase C"],
                    "Post-Fault Max Phase Voltage (pu)": max_post_fault_pu,
                    "Post-Fault Min Phase Voltage (pu)": min_post_fault_pu,
                    "Fault Level":fault_level,
                }

            results_df = pd.concat([results_df, pd.DataFrame([new_row])], ignore_index=True)

    results_df.to_csv(output_csv_path)
   
    fig, ax = plt.subplots(nrows=1, ncols=2, figsize=(16, 9), sharex=False)
    plt.subplots_adjust(left=0.08, right=0.95, bottom=0.1, top=0.90, wspace=0.18, hspace=0.28)
    
    plt.title("Phase Voltage Health",loc='center')

    ax[0].grid(True)
    ax[1].grid(True)

    
    if groupby_str_for_plot_title is not None:
        title_suffix = f" ({groupby_str_for_plot_title})"
    else:
        title_suffix = ""

    ax[0].set_title("In-Fault Healthy Phase Voltages"+title_suffix,pad=10)
    data = [
        results_df[results_df["Fault Type"]=="1PhG"]["In-Fault Max Phase Voltage (pu)"],
        results_df[results_df["Fault Type"]=="2PhG"]["In-Fault Max Phase Voltage (pu)"],
        results_df[results_df["Fault Type"]=="PhPh"]["In-Fault Max Phase Voltage (pu)"],
    ]
    ax[0].boxplot(data,vert=True,widths=0.4)
    ax[0].set_xticklabels([f"1PhG \n (n={len(data[0])})",f"2PhG \n (n={len(data[1])})",f"PhPh \n (n={len(data[2])})"])
    ax[0].set_ylabel("Max In-Fault POC Voltage (pu)")


    ax[1].set_title("Post-Fault Phase Voltages"+title_suffix,pad=10)
    data = [
        results_df[results_df["Fault Type"]=="1PhG"]["Post-Fault Max Phase Voltage (pu)"],
        results_df[results_df["Fault Type"]=="2PhG"]["Post-Fault Max Phase Voltage (pu)"],
        results_df[results_df["Fault Type"]=="PhPh"]["Post-Fault Max Phase Voltage (pu)"],
        results_df[results_df["Fault Type"]=="3PhG"]["Post-Fault Max Phase Voltage (pu)"],
    ]
    ax[1].boxplot(data,vert=True,widths=0.4)
    ax[1].set_xticklabels([f"1PhG \n (n={len(data[0])})",f"2PhG \n (n={len(data[1])})",f"PhPh \n (n={len(data[2])})",f"3PhG \n (n={len(data[3])})"])
    ax[1].set_ylabel("Max Post-Fault POC Voltage (pu)")


    plt.savefig(output_png_path)
    plt.clf()