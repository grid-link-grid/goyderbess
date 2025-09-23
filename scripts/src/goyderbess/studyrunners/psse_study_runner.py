import sys
import shutil
from pallet.utils.time_utils import now_str
import numpy as np
from pathlib import Path

PSSBIN_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSBIN"""
PSSPY_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSPY39"""
sys.path.append(PSSBIN_PATH)
sys.path.append(PSSPY_PATH)
import psse34
import psspy
import json

import os, glob
import pandas as pd
from icecream import ic
import cattrs
from pallet.specs import load_specs_from_csv, load_specs_from_xlsx
from pallet.psse import system, statics
from pallet.psse.initialiser.Initialiser import Initialiser
from pallet.psse.case.CaseDefinition import CaseDefinition
from pallet.psse.PssePallet import PssePallet
from gridlink.utils.wan.wan_utils import find_files
from pallet.psse.case.Bus import Bus, BusData
from pallet.psse.vslack_calculator import calc_vslack
from pallet.utils.alternative_dyntools import out_to_df




def get_vslacks(specification,MODEL_DIR,slack_bus_num):

    case_dir = find_files(MODEL_DIR,".sav")[0]
    savdef_path = find_files(MODEL_DIR,".savdef")[0]
    initdef_path = find_files(MODEL_DIR,".initdef")[0]

    with open(savdef_path, 'r') as f:
        case_definition_raw = json.load(f)
    savdef = cattrs.structure(case_definition_raw, CaseDefinition)

    with open(initdef_path, 'r') as f:
        initialiser_definition_raw = json.load(f)
    initialiser = cattrs.structure(initialiser_definition_raw, Initialiser)
    vslacks = []
    if "Vslack_pu_psse" in specification.columns:
        specification = specification[specification["Vslack_pu_psse"].astype(str).str.contains(r"\$VSLACK")]
        for index, scenario in specification.iterrows():
            system.init(buses=80000)
            statics.load_sav(case_dir)
            initialiser.run(scenario=scenario, case_definition=savdef, locked_taps=False)
            slack_bus_num = slack_bus_num
            ierr, V_PU = psspy.busdat(slack_bus_num,'PU')
            vslacks.append(V_PU)
        dict = {'Vslack_pu_psse': vslacks}
        vslack_df = pd.DataFrame(dict)
        specification['Vslack_pu_psse'] = [
        text.replace("$VSLACK", str(value)) for text, value in zip(specification["Vslack_pu_psse"], vslack_df["Vslack_pu_psse"])
        ]
    else:
         print('no vgrid tests')
    return specification



INF_INIT_GRID_SCR = 3.68
INF_INIT_GRID_X2R = 5
INF_INIT_GRID_FL = 920

#Compatible with pallet version 3.15.1
def run_psse_studies(
        spec,
        plotter: object,
        MODEL_DIR,
        RESULTS_DIR,
  ):


    system.init()
    system.suppress_convergence_monitor()


    # Prepare for infinite SCR studies
    spec["Is_Infinite"] = spec["Grid_SCR"] == "ZERO_IMP"
    spec.loc[spec["Is_Infinite"], "Grid_SCR"] = INF_INIT_GRID_SCR
    spec.loc[spec["Is_Infinite"], "Grid_X2R_sig"] = INF_INIT_GRID_X2R
    spec.loc[spec["Is_Infinite"], "Grid_FL_MVA_sig"] = INF_INIT_GRID_FL

    for is_infinite in [False, True]:
        spec_subset = spec[spec["Is_Infinite"] == is_infinite]
        spec_subset.to_csv("export.csv")
        if len(spec_subset) == 0:
            continue

        p = PssePallet()
        try:
                
             
            p.set_model_directory(MODEL_DIR)
            # if i==0:
            #     spec = spec[(spec["Grid_SCR"] == float("INF"))]
            #     p.set_pre_initialisation_static_callback(callback=pre_init_callback())
            # if i==1:
            #     spec = spec[(spec["Grid_SCR"] != float("INF"))]
            p.set_results_dir(RESULTS_DIR)
            p.set_subfolder_column("Category")
            p.set_result_basename_column("File_Name")
            p.set_end_time_column("Post_Init_Duration_s")
            p.set_command_column("PSSE Commands")
            p.set_playback_columns(
                    fslack_hz_column='Fslack_Hz_sig',
                    vslack_pu_column='Vslack_pu_psse',
                    )
            p.set_dynamic_solution_parameters(
                max_iterations = 1000,
                acceleration_factor = 0.1,
                convergence_tolerance = 0.0005,
                time_step = 0.001,
                frequency_filter = 0.016
            )
        


            def post_init_callback():
                statics.save_case(os.path.join(RESULTS_DIR,"pre_callback_case.sav"))
                assert p._case_definition is not None
                machines = p._case_definition.machines
                if machines is not None:
                        for machine in machines:
                            machine.update_or_create(
                                ireg=machine.bus,
                                vsched=Bus(machine.bus).get_data(BusData.PU),
                            )
                grid_branch = p._case_definition.get_branch("grid").update_or_create(r=0.0,x=0.0)
                poc_voltage = p._case_definition.get_bus("POI").get_data(BusData.PU)
                poc_angle = p._case_definition.get_bus("POI").get_data(BusData.ANGLE)
                p._case_definition.get_machine("slack").update_or_create(
                                vsched=poc_voltage,
                            )
                statics.solve()
                # solved = statics.solved()
                statics.save_case(os.path.join(RESULTS_DIR,"post_callback_case.sav"))

            if is_infinite:    
                p.set_post_initialisation_static_callback(callback=post_init_callback)
            
            spec_subset.to_csv("running_spec"+"_"+str(int(is_infinite))+".csv")
            p.set_spec(spec_subset)
            p.disable_frequency_dependence()
            p.add_plotter(plotter)
            p.set_steps_per_write_column_name("Steps_Per_Write")
            p.run()
            spec_subset.to_csv(f"export_{int(is_infinite)}.csv")
            
        except Exception as e:
            # traceback.print_exc()
            print(e)
            statics.save_case(os.path.join(RESULTS_DIR,"case.sav"))
        

    return 
