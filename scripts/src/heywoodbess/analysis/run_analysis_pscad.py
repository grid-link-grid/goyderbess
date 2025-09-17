import sys
import os
import json
#os.environ["POWERKIT_LICENSE_KEY"] = "396119-1D6701-6F484D-30E154-E30B6A-V3"
import pandas as pd 
import numpy as np
import math
import glob
import traceback
import ast
from pallet.pscad.PscadPallet import PscadPallet
from pallet.pscad.launching import launch_pscad
from pallet.pscad.validation import ValidatorBehaviour
from pallet import load_specs_from_csv, load_specs_from_xlsx
from pallet import now_str



# extension = ".psout"


from pallet.utils.time_utils import now_str
from gridlink.utils.spec_utils import sort_grouped_paths_by_spec
from heywoodbess.analysis.clause_analysis.s52511_dpdf_analysis import produce_s52511_dpdf_outputs
from heywoodbess.analysis.clause_analysis.s5254_cuo_analysis import produce_cuo_summary_outputs
from heywoodbess.analysis.clause_analysis.s52513_vdroop_analysis import produce_vdroop_outputs
from heywoodbess.analysis.clause_analysis.s52513_disturbance_analysis import produce_disturbance_analysis_outputs
from heywoodbess.analysis.clause_analysis.s5255_iq_analysis import produce_s5255_iq_outputs

from heywoodbess.analysis.clause_analysis.cumulative_rt_characteristics import produce_cumulative_ride_through_outputs
from heywoodbess.analysis.clause_analysis.s5255_rise_settle_recovery_analysis import produce_s5255_rise_settle_recovery_curve
from heywoodbess.analysis.clause_analysis.s5258_disturbance_analysis import p_reduction_table
from heywoodbess.analysis.clause_analysis.s52513_ort_analysis import produce_ort_plots
from heywoodbess.analysis.clause_analysis.s528_max_fault_current_analysis import produce_s528_max_fault_current_outputs
from heywoodbess.analysis.clause_analysis.s5255_mfrt_analysis import produce_s5255_mfrt_outputs
from heywoodbess.analysis.clause_analysis.s5255_phase_health_analysis import conduct_phase_voltage_analysis





def run_analysis_pscad(
            x86:bool,
            extension:str,
            ANALYSIS_TO_RUN,
            DPDF_CHARACTERISTIC_POINTS,
            VDROOP_CHARACTERISTIC_POINTS,
            LVRT_THRESHOLDS,
            HVRT_THRESHOLDS,
            DMAT_INPUTS_DIR,
            CSR_INPUTS_DIR,
            OUTPUTS_DIR
                       ):

    # Spec keys
    BALANCED_FAULT_COMMAND_SPEC_KEY = "Fault_Timing_Signal_sig"
    FAULT_TYPE_COMMAND_SPEC_KEY = "Fault_Type_sig"
    INV_ERR_MSG_SPEC_KEY = "INV_OPERATION_STATE"
    INV_OPERATION_STATE_SPEC_KEY = "INV_ERROR_REGISTER"
    VPOC_DISTURBANCE_COMMAND_SPEC_KEY = "Vpoc_disturbance_pu_sig"
    PFREF_PU_COMMAND_SPEC_KEY = "PFref_sig"  #Added SPEC key --> from GESF
    INV_IQ_PU_SIGNAL_NAME = "Id_inv_pos_pu"

    INIT_TIME_SPEC_KEY = "TIME_Full_Init_Time_sec"
    FSLACK_HZ_COMMAND_SPEC_KEY = "Fslack_Hz_sig"
    VSLACK_PU_COMMAND_SPEC_KEY = "Vslack_pu_sig"
    VPOC_PU_COMMAND_SPEC_KEY = "Vpoc_pu_sig"
    VREF_PU_COMMAND_SPEC_KEY = "Vref_pu_sig"
    QREF_PU_COMMAND_SPEC_KEY = "Qref_MVAr_sig"
    FAULT_TIMING_SIGNAL_COMMAND_SPEC_KEY = "Fault_Timing_Signal_sig"
    TOV_TIMING_SIGNAL_COMMAND_SPEC_KEY="TOV_Timing_Signal_sig"
    FAULT_TYPE_COMMAND_SPEC_KEY="Fault_Type_sig"
    GRID_FAULT_LEVEL_SPEC_KEY = "Grid_FL_MVA_sig"
    ORT_OSC_FREQ_KEY="Vslack_osc_Hz_sig"
    ORT_OSC_AMP_KEY="Vslack_osc_amplitude_sig"
    ORT_OSC_ANGLE_KEY="Vslack_osc_phase_deg_sig"

    # Signals
    PREF_MW_SIGNAL_NAME = "Pref_MW"
    PPOC_MW_SIGNAL_NAME = "PLANT_P_HV"
    FPOC_HZ_SIGNAL_NAME = "Fpoc_Hz"
    QPOC_MVAR_SIGNAL_NAME = "PLANT_Q_HV"
    QREF_MVAR_SIGNAL_NAME = "PwrRtSpntTot_MVAr"
    VPOC_RMS_PU_SIGNAL_NAME = "PLANT_V_HV"
    VPOC_RMS_PU_SIGNAL_NAME_2 = "PLANT_V_AN"
    VREF_PU_SIGNAL_NAME = "Vref_pu"
    INV1_VTERM_RMS_PU_SIGNAL_NAME = "VMEAS_LCL" #Added signal 25/02
    PERR_MW_SIGNAL_NAME = "poc_p_err_MW"
    QERR_MVAR_SIGNAL_NAME = "poc_q_err_MVAr"
    VERR_ADJ_PU_SIGNAL_NAME = "poc_v_err_adj_pu"
    #INV_IQ_POS_SIGNAL_NAME = "Iq_inv_pos_pu"
    INV_IQ_NEG_SIGNAL_NAME = "PCU1_Iq_inv_neg_pu"
    INV1_IQ_POS_SIGNAL_NAME = "Iq_pos_LV" #Added signal   25/02
    POC_IRMS_PU_SIGNAL_NAME = "Irms_HV_pu"
    POC_IRMS_KA_SIGNAL_NAME = "Irms_HV"
    POC_IQ_POS_SIGNAL_NAME = "Iq_pos_poc" #Added Signal    25/02
    POC_VRMS_PHA_SIGNAL_NAME="V_hv_rms_phase_A"
    POC_VRMS_PHB_SIGNAL_NAME="V_hv_rms_phase_B"
    POC_VRMS_PHC_SIGNAL_NAME="V_hv_rms_phase_C"
    FRT_FLAG_SIGNAL_NAME = "VRT_FLAG"
    FRT_FLAG_ACTIVE_VALUES = [1,2,3]
    ERROR_SIGNAL_NAME = "ERROR"

    # Characteristics
    DIQDV_CHARACTERISTIC_POINTS = [(0,1.0),(0.6,1),(0.85,0),(1.15,0),(1.32,-1),(2.0,-1.0)] #AAS
    HFRT_CHARACTERISTIC_POINTS = [(0, 52.001), (600, 52), (600, 50.5001), (800, 50.5)] #AAS
    LFRT_CHARACTERISTIC_POINTS = [(0, 46.999), (120, 47), (120, 47.999), (600, 48), (600, 49.499), (800, 49.5)] #AAS
    HVRT_CHARACTERISTIC_POINTS = [(0, 1.3501), (0.02, 1.35), (0.02, 1.301), (0.2, 1.3), (0.2, 1.2501), (2.0, 1.25), (2.0, 1.201), (20, 1.2), (20, 1.151), (1200, 1.15), (1200, 1.1101), (1500, 1.1)]  #AAS
    LVRT_CHARACTERISTIC_POINTS = [(0, 0.699), (2, 0.7), (2, 0.799), (10, 0.8), (10, 0.899), (1500, 0.9)] #AAS
    DIQDV_INV_CHARACTERISTIC_POINTS = [(0.5,1),(0.685,1.0),(0.85,0),(1.18,0),(1.345,-1),(1.5,-1)]  #Added characteristic 25/02 - Copied from GESF #AAS
    # Constants
    DIQDV_STARTING_VOLTAGE_OV_PU = 1.2
    DIQDV_STARTING_VOLTAGE_UV_PU = 0.8

    DIQDV_INCLUDED_VOLTAGE_RANGES = [[0,0.9],[1.1,1.5]] #Added Constants  25/02 - Copied from GESF
    DIQDV_INV_INCLUDED_VOLTAGE_RANGES = [[0.0,1],[1.18,1.3]]  #Added Constants  25/02 - Copied from GESF
    DIQDV_VOLTAGE_THRESHOLDS_PER_GPS = [0.9, 1.1] 

    # To be converted into spec keys when data in pscad
    PMAX_MW = 285.0
    PMIN_MW = -285.0

    PSCAD_FAULT_TYPE_ENUMS = {
        0: "No Fault",
        1: "1PhG",
        2: "1PhG",
        3: "1PhG",
        4: "2PhG",
        5: "2PhG",
        6: "2PhG",
        7: "3PhG",
        8: "PhPh",
        9: "PhPh",
        10: "PhPh",
        11: "3Ph",
        }
    
    PSCAD_HEALTHY_PHASE_ENUMS = {
        0: ["Phase A","Phase B","Phase C"],
        1: ["Phase B", "Phase C"],
        2: ["Phase A", "Phase C"],
        3: ["Phase A","Phase B"],
        4: ["Phase C"],
        5: ["Phase B"],
        6: ["Phase A"],
        7: [],
        8: ["Phase C"],
        9: ["Phase B"],
        10: ["Phase A"],
        11: [],
    }

    errors = []
    if not os.path.exists(OUTPUTS_DIR):
        os.makedirs(OUTPUTS_DIR)


    if "dP/df characteristic" in ANALYSIS_TO_RUN:
        out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR Fgrid Steps") + "/*"+extension)
        #print(out_paths)
        filtered_paths = [s for s in out_paths if "52511" in s]
        try:
            produce_s52511_dpdf_outputs(
                    psout_paths = filtered_paths,
                    output_png_path=os.path.join(OUTPUTS_DIR, "dpdf_characteristic.png"),
                    output_csv_path=os.path.join(OUTPUTS_DIR, "dpdf_results.csv"),
                    init_time_spec_key = INIT_TIME_SPEC_KEY,
                    pref_mw_signal_name=PREF_MW_SIGNAL_NAME,
                    ppoc_mw_signal_name=PPOC_MW_SIGNAL_NAME,
                    fpoc_hz_signal_name=FPOC_HZ_SIGNAL_NAME,
                    fslack_hz_command_spec_key = FSLACK_HZ_COMMAND_SPEC_KEY,
                    dpdf_characteristic_points=DPDF_CHARACTERISTIC_POINTS,
                    df_manipulation_fn=None,
                    pmax_mw=PMAX_MW,
                    pmin_mw=PMIN_MW,
                    x86=x86
                    )
        except Exception as e:
            print("Failed on:     dP/df characteristic")
            traceback.print_exc()
            print(e)
            errors.append(f"dP/df characteristic: {e}")

    if "CUO" in ANALYSIS_TO_RUN:

        out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR CUO") + "/*"+extension)
        print(out_paths)
        # filtered_paths = [s for s in out_paths if "CUO" in s]
        try: 
            produce_cuo_summary_outputs(
                    psout_paths = out_paths,
                    output_png_path=os.path.join(OUTPUTS_DIR, "cuo_summary.png"),
                    output_csv_path=os.path.join(OUTPUTS_DIR, "cuo_results.csv"),
                    init_time_spec_key = INIT_TIME_SPEC_KEY,
                    vterm_rms_pu_signal_name=INV1_VTERM_RMS_PU_SIGNAL_NAME,
                    vpoc_pu_command_spec_key=VPOC_DISTURBANCE_COMMAND_SPEC_KEY,
                    hvrt_thresholds=HVRT_THRESHOLDS,
                    lvrt_thresholds=LVRT_THRESHOLDS,
                    x86=x86
                    )
        except Exception as e:
            print("Failed on:     CUO")
            traceback.print_exc()
            print(e)
            errors.append(f"CUO: {e}")
            


    if "Vdroop characteristic" in ANALYSIS_TO_RUN:
        out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR Vgrid Steps") + "/*"+extension)
        # filtered_paths = [s for s in out_paths if "Vdroop" in s]
        try:
            produce_vdroop_outputs(
                    psout_paths = out_paths,
                    output_png_path=os.path.join(OUTPUTS_DIR, "vdroop_characteristic.png"),
                    output_csv_path=os.path.join(OUTPUTS_DIR, "vdroop_results.csv"),
                    qpoc_mvar_signal_name=QPOC_MVAR_SIGNAL_NAME,
                    vpoc_rms_pu_signal_name=VPOC_RMS_PU_SIGNAL_NAME,
                    vref_pu_signal_name=VREF_PU_SIGNAL_NAME,
                    vpoc_pu_command_spec_key=VSLACK_PU_COMMAND_SPEC_KEY,
                    vdroop_characteristic_points=VDROOP_CHARACTERISTIC_POINTS,
                    init_time_spec_key = INIT_TIME_SPEC_KEY,
                    df_manipulation_fn = None,
                    x86 = x86
                    )
        except Exception as e:
            print("failed on:      Vdroop characteristic")
            print(e)
            traceback.print_exc()    
            errors.append(f"Vdroop characteristic: {e}")        


    if "Vgrid step analysis" in ANALYSIS_TO_RUN:
        out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR Vgrid Steps") + "/*"+extension)
    
        # filtered_paths = [s for s in out_paths if ("Vdroop" in s)]

        saturated_paths=[]
        unsaturated_paths = []

        for path in out_paths:
            json_path = path.replace(extension, ".json")
            with open(json_path,'r') as f: 
                spec = json.load(f)
            if "Saturated_Status" in spec:
                if spec["Saturated_Status"] == "YES":
                    saturated_paths.append(path)
                if spec["Saturated_Status"] == "NO":
                    unsaturated_paths.append(path)

        try:
            # if len(saturated_paths)!=0:
            produce_disturbance_analysis_outputs(
                    psout_paths = out_paths,
                    output_png_path=os.path.join(OUTPUTS_DIR, "Vgrid droop step disturbance tracking saturated.png"),
                    output_csv_path=os.path.join(OUTPUTS_DIR, "Vgrid droop step disturbance tracking saturated.csv"),
                    disturbance_command_spec_key=VSLACK_PU_COMMAND_SPEC_KEY,
                    ppoc_mw_signal_name=PPOC_MW_SIGNAL_NAME,
                    qpoc_mvar_signal_name=QPOC_MVAR_SIGNAL_NAME,
                    vpoc_pu_signal_name=VPOC_RMS_PU_SIGNAL_NAME_2,
                    perr_mw_signal_name=PERR_MW_SIGNAL_NAME,
                    qerr_mvar_signal_name=QERR_MVAR_SIGNAL_NAME,
                    verr_adj_pu_signal_name=VERR_ADJ_PU_SIGNAL_NAME,
                    rise_settle_title_spec_key="File_Name",
                    pre_step_plot_secs=2.0,
                    pre_error_plot_secs=2.0,
                    init_time_spec_key = INIT_TIME_SPEC_KEY,
                    df_manipulation_fn = None,
                    x86 = x86
                    )
        except Exception as e:
            print("Failed on:     Vgrid step analysis")
            print(e)
            traceback.print_exc()
            errors.append(f"Vgrid step analysis (sat): {e}")
        
        try:
            if len(unsaturated_paths)!=0:
                produce_disturbance_analysis_outputs(
                        psout_paths = unsaturated_paths,
                        output_png_path=os.path.join(OUTPUTS_DIR, "Vgrid droop step disturbance tracking unsaturated.png"),
                        output_csv_path=os.path.join(OUTPUTS_DIR, "Vgrid droop step disturbance tracking unsaturated.csv"),
                        disturbance_command_spec_key=VSLACK_PU_COMMAND_SPEC_KEY,
                        ppoc_mw_signal_name=PPOC_MW_SIGNAL_NAME,
                        qpoc_mvar_signal_name=QPOC_MVAR_SIGNAL_NAME,
                        vpoc_pu_signal_name=VPOC_RMS_PU_SIGNAL_NAME_2,
                        perr_mw_signal_name=PERR_MW_SIGNAL_NAME,
                        qerr_mvar_signal_name=QERR_MVAR_SIGNAL_NAME,
                        verr_adj_pu_signal_name=VERR_ADJ_PU_SIGNAL_NAME,
                        rise_settle_title_spec_key="File_Name",
                        pre_step_plot_secs=2.0,
                        pre_error_plot_secs=2.0,
                        init_time_spec_key = INIT_TIME_SPEC_KEY,
                        df_manipulation_fn = None,
                        x86 = x86
                        )
        except Exception as e:
            print("Failed on:     Vgrid step analysis")
            print(e)
            traceback.print_exc()
            errors.append(f"Vgrid step analysis (unsat): {e}")
        
    #5_2_5_13 -- END        


    if "Vref step analysis" in ANALYSIS_TO_RUN:
        out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR Vref Steps") + "/*"+extension)
        filtered_paths = [s for s in out_paths if "CSR" in s]

        print(filtered_paths)

        saturated_paths=[]
        unsaturated_paths = []

        for path in filtered_paths:
            json_path = path.replace(extension, ".json")
            with open(json_path,'r') as f: 
                spec = json.load(f)
            if "Saturated_Status" in spec:
                if spec["Saturated_Status"] == "YES":
                    saturated_paths.append(path)
                if spec["Saturated_Status"] == "NO":
                    unsaturated_paths.append(path)



        try:
            # if len(saturated_paths)!=0:
            produce_disturbance_analysis_outputs(
                    psout_paths = filtered_paths,
                    output_png_path=os.path.join(OUTPUTS_DIR, "Vref step disturbance tracking saturated.png"),
                    output_csv_path=os.path.join(OUTPUTS_DIR, "Vref step disturbance tracking saturated.csv"),
                    disturbance_command_spec_key=VREF_PU_COMMAND_SPEC_KEY,
                    ppoc_mw_signal_name=PPOC_MW_SIGNAL_NAME,
                    qpoc_mvar_signal_name=QPOC_MVAR_SIGNAL_NAME,
                    vpoc_pu_signal_name=VPOC_RMS_PU_SIGNAL_NAME_2,
                    perr_mw_signal_name=PERR_MW_SIGNAL_NAME,
                    qerr_mvar_signal_name=QERR_MVAR_SIGNAL_NAME,
                    verr_adj_pu_signal_name=VERR_ADJ_PU_SIGNAL_NAME,
                    rise_settle_title_spec_key="File_Name",
                    init_time_spec_key = INIT_TIME_SPEC_KEY,
                    df_manipulation_fn = None,
                    x86 = x86
                    )
        except Exception as e:
            print("Failed on:    Vref step analysis")
            print(e)
            traceback.print_exc()
            errors.append(f"Vref step analysis (sat): {e}")

        try:     
            if len(unsaturated_paths)!=0:
                produce_disturbance_analysis_outputs(
                        psout_paths = unsaturated_paths,
                        output_png_path=os.path.join(OUTPUTS_DIR, "Vref step disturbance tracking unsaturated.png"),
                        output_csv_path=os.path.join(OUTPUTS_DIR, "Vref step disturbance tracking unsaturated.csv"),
                        disturbance_command_spec_key=VREF_PU_COMMAND_SPEC_KEY,
                        ppoc_mw_signal_name=PPOC_MW_SIGNAL_NAME,
                        qpoc_mvar_signal_name=QPOC_MVAR_SIGNAL_NAME,
                        vpoc_pu_signal_name=VPOC_RMS_PU_SIGNAL_NAME_2,
                        perr_mw_signal_name=PERR_MW_SIGNAL_NAME,
                        qerr_mvar_signal_name=QERR_MVAR_SIGNAL_NAME,
                        verr_adj_pu_signal_name=VERR_ADJ_PU_SIGNAL_NAME,
                        rise_settle_title_spec_key="File_Name",
                        init_time_spec_key = INIT_TIME_SPEC_KEY,
                        df_manipulation_fn = None,
                        x86 = x86
                        )
        except Exception as e:
            print("Failed on:     Vref step analysis")
            print(e)
            traceback.print_exc()
            errors.append(f"Vref step analysis (sat): {e}")


    if "Qref step analysis" in ANALYSIS_TO_RUN:
        out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR Qref Steps") + "/*"+extension)
        #print(out_paths)
        filtered_paths = [s for s in out_paths if "CSR" in s]
        # saturated_paths=[]
        # unsaturated_paths = []

        # for path in filtered_paths:
        #     json_path = path.replace(extension, ".json")
        #     with open(json_path,'r') as f: 
        #         spec = json.load(f)
        #     if "Saturated_Status" in spec:
        #         if spec["Saturated_Status"] == "YES":
        #             saturated_paths.append(path)
        #         if spec["Saturated_Status"] == "NO":
        #             unsaturated_paths.append(path)
        try:

            # if len(saturated_paths)!=0:
            produce_disturbance_analysis_outputs(
                    psout_paths = filtered_paths,
                    output_png_path=os.path.join(OUTPUTS_DIR, "Qref step disturbance tracking saturated.png"),
                    output_csv_path=os.path.join(OUTPUTS_DIR, "Qref step disturbance tracking saturated.csv"),
                    disturbance_command_spec_key=QREF_PU_COMMAND_SPEC_KEY,
                    ppoc_mw_signal_name=PPOC_MW_SIGNAL_NAME,
                    qpoc_mvar_signal_name=QPOC_MVAR_SIGNAL_NAME,
                    vpoc_pu_signal_name=VPOC_RMS_PU_SIGNAL_NAME_2,
                    perr_mw_signal_name=PERR_MW_SIGNAL_NAME,
                    qerr_mvar_signal_name=QERR_MVAR_SIGNAL_NAME,
                    verr_adj_pu_signal_name=VERR_ADJ_PU_SIGNAL_NAME,
                    rise_settle_title_spec_key="File_Name",
                    init_time_spec_key = INIT_TIME_SPEC_KEY,
                    df_manipulation_fn = None,
                    x86 = x86
                    )
        except Exception as e:
            print("Failed on:     Qref step analysis")
            print(e)
            traceback.print_exc()
            errors.append(f"Qref step analysis (sat): {e}")

        # try:
        #     if len(unsaturated_paths)!=0:
        #         produce_disturbance_analysis_outputs(
        #                 psout_paths = unsaturated_paths,
        #                 output_png_path=os.path.join(OUTPUTS_DIR, "Qref step disturbance tracking unsaturated.png"),
        #                 output_csv_path=os.path.join(OUTPUTS_DIR, "Qref step disturbance tracking unsaturated.csv"),
        #                 disturbance_command_spec_key=QREF_PU_COMMAND_SPEC_KEY,
        #                 ppoc_mw_signal_name=PPOC_MW_SIGNAL_NAME,
        #                 qpoc_mvar_signal_name=QPOC_MVAR_SIGNAL_NAME,
        #                 vpoc_pu_signal_name=VPOC_RMS_PU_SIGNAL_NAME_2,
        #                 perr_mw_signal_name=PERR_MW_SIGNAL_NAME,
        #                 qerr_mvar_signal_name=QERR_MVAR_SIGNAL_NAME,
        #                 verr_adj_pu_signal_name=VERR_ADJ_PU_SIGNAL_NAME,
        #                 rise_settle_title_spec_key="File_Name",
        #                 init_time_spec_key = INIT_TIME_SPEC_KEY,
        #                 df_manipulation_fn = None,
        #                 x86 = x86
        #                 )
        # except Exception as e:
        #     print("Failed on:     Qref step analysis")
        #     print(e)
        #     traceback.print_exc()
        #     errors.append(f"Qref step analysis (unsat): {e}")



    if "PFref step analysis" in ANALYSIS_TO_RUN:
        out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR PFref Steps") + "/*"+extension)
        #print(out_paths)
        filtered_paths = [s for s in out_paths if "CSR" in s]

        # saturated_paths=[]
        # unsaturated_paths = []

        # for path in filtered_paths:
        #     json_path = path.replace(extension, ".json")
        #     with open(json_path,'r') as f: 
        #         spec = json.load(f)
        #     if "Saturated_Status" in spec:
        #         if spec["Saturated_Status"] == "YES":
        #             saturated_paths.append(path)
        #         if spec["Saturated_Status"] == "NO":
        #             unsaturated_paths.append(path)      

        try:

            # if len(saturated_paths)!=0:
            produce_disturbance_analysis_outputs(
                    psout_paths = filtered_paths,
                    output_png_path=os.path.join(OUTPUTS_DIR, "PFref step disturbance tracking saturated.png"),
                    output_csv_path=os.path.join(OUTPUTS_DIR, "PFref step disturbance tracking saturated.csv"),
                    disturbance_command_spec_key=PFREF_PU_COMMAND_SPEC_KEY,
                    ppoc_mw_signal_name=PPOC_MW_SIGNAL_NAME,
                    qpoc_mvar_signal_name=QPOC_MVAR_SIGNAL_NAME,
                    vpoc_pu_signal_name=VPOC_RMS_PU_SIGNAL_NAME_2,
                    perr_mw_signal_name=PERR_MW_SIGNAL_NAME,
                    qerr_mvar_signal_name=QERR_MVAR_SIGNAL_NAME,
                    verr_adj_pu_signal_name=VERR_ADJ_PU_SIGNAL_NAME,
                    rise_settle_title_spec_key="File_Name",
                    init_time_spec_key = INIT_TIME_SPEC_KEY,
                    df_manipulation_fn = None,
                    x86 = x86
                    )

        except Exception as e:
            print("Failed on:    PFref step analysis")
            print(e)
            traceback.print_exc()
            errors.append(f"PFref step analysis (sat): {e}")
        # try:
        #     if len(unsaturated_paths)!=0:
        #         produce_disturbance_analysis_outputs(
        #                 psout_paths = unsaturated_paths,
        #                 output_png_path=os.path.join(OUTPUTS_DIR, "PFref step disturbance tracking unsaturated.png"),
        #                 output_csv_path=os.path.join(OUTPUTS_DIR, "PFref step disturbance tracking unsaturated.csv"),
        #                 disturbance_command_spec_key=PFREF_PU_COMMAND_SPEC_KEY,
        #                 ppoc_mw_signal_name=PPOC_MW_SIGNAL_NAME,
        #                 qpoc_mvar_signal_name=QPOC_MVAR_SIGNAL_NAME,
        #                 vpoc_pu_signal_name=VPOC_RMS_PU_SIGNAL_NAME_2,
        #                 perr_mw_signal_name=PERR_MW_SIGNAL_NAME,
        #                 qerr_mvar_signal_name=QERR_MVAR_SIGNAL_NAME,
        #                 verr_adj_pu_signal_name=VERR_ADJ_PU_SIGNAL_NAME,
        #                 rise_settle_title_spec_key="File_Name",
        #                 init_time_spec_key = INIT_TIME_SPEC_KEY,
        #                 df_manipulation_fn = None,
        #                 x86 = x86
        #                 )
        # except Exception as e:
        #     print("Failed on:    PFref step analysis")
        #     print(e)
        #     traceback.print_exc()
        #     errors.append(f"PFref step analysis (unsat): {e}")

    if "diq/dV characteristic" in ANALYSIS_TO_RUN:

        balanced_fault_out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR Balanced Faults") + "/*"+extension)
        unbalanced_fault_out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR Unbalanced Faults") + "/*"+extension)
        fault_out_paths = balanced_fault_out_paths + unbalanced_fault_out_paths
        tov_out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR TOV") + "/*"+extension)[:]

        try:
            produce_s5255_iq_outputs(
                    fault_psout_paths=fault_out_paths,
                    tov_psout_paths=tov_out_paths,
                    output_png_path=os.path.join(OUTPUTS_DIR, "diqdv_characteristic.png"),
                    output_csv_path=os.path.join(OUTPUTS_DIR, "diqdv_results.csv"),
                    groupby_str_for_plot_titles=None,
                    diqdv_characteristic_points=DIQDV_CHARACTERISTIC_POINTS,
                    init_time_spec_key=INIT_TIME_SPEC_KEY,
                    fault_timing_signal_command_spec_key=FAULT_TIMING_SIGNAL_COMMAND_SPEC_KEY,
                    tov_timing_signal_command_spec_key=TOV_TIMING_SIGNAL_COMMAND_SPEC_KEY,
                    iq_pos_seq_signal_name=POC_IQ_POS_SIGNAL_NAME,
                    iq_neg_seq_signal_name="NA",
                    v_signal_name=VPOC_RMS_PU_SIGNAL_NAME,
                    saturation_current_signal_name=POC_IRMS_PU_SIGNAL_NAME,
                    included_voltage_ranges=DIQDV_INCLUDED_VOLTAGE_RANGES,
                    gps_voltage_thresholds=DIQDV_VOLTAGE_THRESHOLDS_PER_GPS,
                    y_label=u'Change in Reactive Current, Δ$I_{q}$ (pu)',
                    x_label=u'Point of Connection Voltage, $V_{fault,settled}$ (pu)',
                    df_manipulation_fn = None,
                    x86 = x86
                    )
            
            produce_s5255_iq_outputs(
                    fault_psout_paths=fault_out_paths,
                    tov_psout_paths=tov_out_paths,
                    output_png_path=os.path.join(OUTPUTS_DIR, "inv1_diqdv_characteristic.png"),
                    output_csv_path=os.path.join(OUTPUTS_DIR, "inv1_diqdv_results.csv"),
                    groupby_str_for_plot_titles=None,
                    diqdv_characteristic_points=DIQDV_INV_CHARACTERISTIC_POINTS,
                    init_time_spec_key=INIT_TIME_SPEC_KEY,
                    fault_timing_signal_command_spec_key=FAULT_TIMING_SIGNAL_COMMAND_SPEC_KEY,
                    tov_timing_signal_command_spec_key=TOV_TIMING_SIGNAL_COMMAND_SPEC_KEY,
                    iq_pos_seq_signal_name=INV1_IQ_POS_SIGNAL_NAME,
                    iq_neg_seq_signal_name="NA",
                    v_signal_name=INV1_VTERM_RMS_PU_SIGNAL_NAME,
                    saturation_current_signal_name=POC_IRMS_PU_SIGNAL_NAME,
                    included_voltage_ranges=DIQDV_INV_INCLUDED_VOLTAGE_RANGES,
                    y_label=u'Change in Reactive Current, Δ$I_{q}$ (pu)',
                    x_label=u'Terminal Voltage, $V_{fault,settled}$ (pu)',
                    df_manipulation_fn = None,
                    x86 = x86
                    )
        except Exception as e:
            print("Failed on:   diq/dV characteristic")
            print(e)
            traceback.print_exc()
            errors.append(f"diq/dV characteristic: {e}")



    if "Frequency ride-through characteristic" in ANALYSIS_TO_RUN:
        out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR F Withstand") + "/*"+extension)[:]


        try:
            produce_cumulative_ride_through_outputs(
                    psout_paths = out_paths,
                    output_png_path=os.path.join(OUTPUTS_DIR, "s5253_ride_through_results.png"),
                    output_csv_path=os.path.join(OUTPUTS_DIR, "s5253_ride_through_results.csv"),
                    step_size=0.05,
                    init_time_spec_key=INIT_TIME_SPEC_KEY,
                    measurement_signal_name=FPOC_HZ_SIGNAL_NAME,
                    high_characteristic_points=HFRT_CHARACTERISTIC_POINTS,
                    low_characteristic_points=LFRT_CHARACTERISTIC_POINTS,
                    #error_signal_name=ERROR_SIGNAL_NAME,
                    df_manipulation_fn=None,
                    title="Frequency ride-through",
                    y_label="Frequency (Hz)",
                    x86 = x86,
                    )
        except Exception as e:
            print("Failed on:   Frequency ride-through characteristic")
            print(e)
            traceback.print_exc()
            errors.append(f"Frequency ride-through characteristic: {e}")
        

    if "Voltage ride-through characteristic" in ANALYSIS_TO_RUN:

        out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR V Withstand") + "/*"+extension)[:]
        filtered_paths = []
        for path in out_paths:
            json_path = path.replace(extension, ".json")
            with open(json_path,'r') as f:
                spec = json.load(f)
            if "Envelope" in spec:
                if spec["Envelope"] == "YES":
                    filtered_paths.append(path)


        try:
            produce_cumulative_ride_through_outputs(
                    psout_paths = filtered_paths,
                    output_png_path=os.path.join(OUTPUTS_DIR, "s5254_ride_through_results.png"),
                    output_csv_path=os.path.join(OUTPUTS_DIR, "s5254_ride_through_results.csv"),
                    step_size=0.001,
                    init_time_spec_key=INIT_TIME_SPEC_KEY,
                    measurement_signal_name=VPOC_RMS_PU_SIGNAL_NAME,
                    high_characteristic_points=HVRT_CHARACTERISTIC_POINTS,
                    low_characteristic_points=LVRT_CHARACTERISTIC_POINTS,
                    #error_signal_name=ERROR_SIGNAL_NAME,
                    df_manipulation_fn=None,
                    title="Voltage ride-through",
                    y_label="Voltage (pu)",
                    x_axis_min=0.0,
                    x_axis_max=1215.0,
                    x86 = x86,
                        )
            print("Completed Voltage ride through characteristic")

        except Exception as e:
            print("Failed on:   Voltage ride-through characteristic")
            print(e)
            traceback.print_exc()
            errors.append(f"Voltage ride-through characteristic: {e}")

    if "IQ Rise Settle & P Recovery Curve" in ANALYSIS_TO_RUN:

        balanced_fault_out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR Balanced Faults") + "/*"+extension)
        unbalanced_fault_out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR Unbalanced Faults") + "/*"+extension)
        tov_out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR TOV") + "/*"+extension)[:]
        fault_out_paths = balanced_fault_out_paths + unbalanced_fault_out_paths
        try: 
            produce_s5255_rise_settle_recovery_curve(
                    psout_paths=fault_out_paths,
                    output_dir=os.path.join(OUTPUTS_DIR,"S5255 POC Faults Rise Settle Recovery Analysis"),
                    disturbance_command_spec_key=FAULT_TIMING_SIGNAL_COMMAND_SPEC_KEY,
                    vpoc_pu_signal_name=VPOC_RMS_PU_SIGNAL_NAME,
                    ppoc_mw_signal_name=PPOC_MW_SIGNAL_NAME,
                    iq_pu_signal_name=POC_IQ_POS_SIGNAL_NAME,
                    frt_flag_signal_name=FRT_FLAG_SIGNAL_NAME,
                    frt_flag_active_values=FRT_FLAG_ACTIVE_VALUES,
                    rise_settle_title_spec_key="File_Name",
                    pre_step_plot_secs=2.0,
                    df_manipulation_fn = None,
                    x86 = x86
                )
            produce_s5255_rise_settle_recovery_curve(
                    psout_paths=tov_out_paths,
                    output_dir=os.path.join(OUTPUTS_DIR,"S5255 POC TOV Rise Settle Recovery Analysis"),
                    disturbance_command_spec_key=TOV_TIMING_SIGNAL_COMMAND_SPEC_KEY,
                    vpoc_pu_signal_name=VPOC_RMS_PU_SIGNAL_NAME,
                    ppoc_mw_signal_name=PPOC_MW_SIGNAL_NAME,
                    iq_pu_signal_name=POC_IQ_POS_SIGNAL_NAME,
                    rise_settle_title_spec_key="File_Name",
                    pre_step_plot_secs=2.0,
                    df_manipulation_fn = None,
                    x86 = x86
                )
        except Exception as e:
            print("Failed on:   IQ Rise Settle & P Recovery Curve")
            print(e)
            traceback.print_exc()
            errors.append(f"IQ Rise Settle & P Recovery Curve: {e}")



    if "S5258 Active Power Reduction" in ANALYSIS_TO_RUN:
        
        out_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR Fgrid Steps") + "/*"+extension)
        filtered_paths = [path for path in out_paths if "5258" in path]
        if not filtered_paths:
            errors.append(f"S5258 Active Power Reduction (CSR Fgrid Steps): NO OUT PATHS DETECTED")  
        try:
            p_reduction_table(
                psout_paths=filtered_paths,
                freq_channel_key=FPOC_HZ_SIGNAL_NAME,
                active_power_channel_key=PPOC_MW_SIGNAL_NAME,
                output_csv_path=os.path.join(OUTPUTS_DIR, "Active Power Reduction Analysis.csv"),
                p_ramp_down_threshold = 0.5,
                freq_init_ramp_down = 51.0,
                freq_complete_ramp_down = 51.0,
                df_manipulation_fn = None,
                pmax=PMAX_MW,
                Discharge_Cases_only=True,
                x86 = x86
                )
        except Exception as e:
            print("Failed on:   S5258 Active Power Reduction")
            print(e)
            traceback.print_exc()
            errors.append(f"S5258 Active Power Reduction: {e}")




    if "ORT Analysis" in ANALYSIS_TO_RUN:

        grouped_sorted_paths = sort_grouped_paths_by_spec(
                sort_by_spec_key=ORT_OSC_FREQ_KEY,
                group_by_spec_keys =['Prefix',ORT_OSC_ANGLE_KEY],
                paths=glob.glob(os.path.join(DMAT_INPUTS_DIR, "ORT") + "/*"+extension),
            )


        try:
            for group_name,grouped_paths in grouped_sorted_paths.items():
                group_str = f"{group_name[0]} Osc Phase={group_name[1]} deg"
                print(group_str)
                
                produce_ort_plots(
                    psout_paths=grouped_paths,
                    vslack_osc_freq_spec_key=ORT_OSC_FREQ_KEY,
                    vslack_osc_amp_spec_key=ORT_OSC_AMP_KEY,
                    vpoc_channel_key=VPOC_RMS_PU_SIGNAL_NAME,
                    qpoc_channel_key=QPOC_MVAR_SIGNAL_NAME,
                    output_path=os.path.join(OUTPUTS_DIR, "ORT Summary Plot "+group_str+".png"),
                    periods_per_subplot=3.0
                )
                
               
        except Exception as e:
            print("Failed on:    ORT Analysis")
            print(e)
            traceback.print_exc()
            errors.append(f"ORT Analysis: {e}")


    if "Max fault current" in ANALYSIS_TO_RUN:
        psout_paths = []
        for subdir in ["CSR Balanced Faults", "CSR Unbalanced Faults"]:
            psout_paths.extend(glob.glob(os.path.join(CSR_INPUTS_DIR, subdir) + "/*"+extension))
        
        try:
            produce_s528_max_fault_current_outputs(
                    psout_paths=psout_paths,
                    output_png_path=os.path.join(OUTPUTS_DIR, "fault_current_results.png"),
                    output_csv_path=os.path.join(OUTPUTS_DIR, "fault_current_results.csv"),
                    init_time_spec_key=INIT_TIME_SPEC_KEY,
                    ia_rms_ka_signal_name=POC_IRMS_KA_SIGNAL_NAME,
                    ib_rms_ka_signal_name=POC_IRMS_KA_SIGNAL_NAME,
                    ic_rms_ka_signal_name=POC_IRMS_KA_SIGNAL_NAME
                    )
        except Exception as e:
            print("Failed on:     Max fault current")
            traceback.print_exc()
            errors.append(f"Max fault current: {e}")

    if "MFRT" in ANALYSIS_TO_RUN:
        try:
            produce_s5255_mfrt_outputs(
                    psout_paths = glob.glob(os.path.join(CSR_INPUTS_DIR, "CSR MFRT") + "/*"+extension),
                    output_csv_path=os.path.join(OUTPUTS_DIR, "mfrt_results.csv"),
                    init_time_spec_key=INIT_TIME_SPEC_KEY,
                    disturbance_command_spec_key=FAULT_TIMING_SIGNAL_COMMAND_SPEC_KEY,
                    fault_type_command_spec_key=FAULT_TYPE_COMMAND_SPEC_KEY,
                    err_msg_spec_key=INV_ERR_MSG_SPEC_KEY,
                    operation_state_spec_key=INV_OPERATION_STATE_SPEC_KEY,
                    )
        except Exception as e:
            print("Failed on:     MFRT") 
            print(e)
            traceback.print_exc()
            errors.append(f"MFRT: {e}")

    if "Phase Health Analysis" in ANALYSIS_TO_RUN:

        try:
            psout_paths = [glob.glob(os.path.join(CSR_INPUTS_DIR, subfolder_name) + "/*"+extension) for subfolder_name in ["CSR Unbalanced Faults", "CSR Balanced Faults"]]
            psout_paths = [item for sublist in psout_paths for item in sublist]
            grouped_sorted_paths = sort_grouped_paths_by_spec(
                    sort_by_spec_key="File_Name",
                    group_by_spec_keys =[GRID_FAULT_LEVEL_SPEC_KEY],
                    paths=psout_paths,
            )

            for group_name,grouped_paths in grouped_sorted_paths.items():
                group_str = f"FL={group_name[0]} MVA"
                conduct_phase_voltage_analysis(
                    psout_paths=grouped_paths,
                    output_png_path=os.path.join(OUTPUTS_DIR,"S5255 Phase Voltage Health", f"{group_str} Phase Voltage Health.png"),
                    output_csv_path=os.path.join(OUTPUTS_DIR,"S5255 Phase Voltage Health", f"{group_str} Phase Voltage Health.csv"),
                    groupby_str_for_plot_title=group_str,
                    timing_cmd_spec_key=FAULT_TIMING_SIGNAL_COMMAND_SPEC_KEY,
                    fault_type_cmd_spec_key=FAULT_TYPE_COMMAND_SPEC_KEY,
                    fault_type_enums = PSCAD_FAULT_TYPE_ENUMS,
                    healthy_phase_enums = PSCAD_HEALTHY_PHASE_ENUMS,
                    phA_channel_key=POC_VRMS_PHA_SIGNAL_NAME,
                    phB_channel_key=POC_VRMS_PHB_SIGNAL_NAME,
                    phC_channel_key=POC_VRMS_PHC_SIGNAL_NAME,
                )
        except Exception as e:
            print("Failed on:     Phase Health Analysis")
            print(e)
            traceback.print_exc()
            errors.append(f"Phase Health Analysis: {e}")


    print("\n")
    print("FINISHED ANALYSIS.\n")
    print("All the Errors Encountered:")
    if len(errors)==0:
        print("No errors")
    else:
        for idx,error in enumerate(errors):
            print(f"{idx+1}. {error}\n")


    return