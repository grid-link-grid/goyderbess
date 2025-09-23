import os
import json
import pandas as pd
import numpy as np
from typing import List, Callable, Optional
from tqdm.auto import tqdm
from pallet.pscad.Psout import Psout
from pallet.dsl.ParsedDslSignal import ParsedDslSignal


def produce_s5255_mfrt_outputs(
        psout_paths : List[str],
        output_csv_path : str,
        init_time_spec_key : str,
        disturbance_command_spec_key : str,
        fault_type_command_spec_key : str,
        err_msg_spec_key : Optional[str] = None,
        operation_state_spec_key : Optional[str] = None,
        df_manipulation_fn : Optional[Callable] = None,
        ):

    results_df = pd.DataFrame(columns=[
            "Name",
            "Fault Times (s)",
            "Fault Duration (s)",
            "Fault Types",
            "Minimum Time Between Faults",
            "Number of 3 Phase Faults",
            "Disturbances with Residual Voltage Below 90%",
            "Disturbances with Residual Voltage Between 70-90%",
            "Time Below Vpoc of 90%",
            "Time Integral (pu.s)",
            "Inverter Operation States",
            "Inverter Error Messages",
        ])


    for i in tqdm(range(len(psout_paths)), desc="MFRT Analysis"):
        psout_path = psout_paths[i]
        json_path = psout_path.split(".psout")[0] + ".json"

        with open(json_path, 'r') as f:
            spec = json.load(f)

        init_time_sec = float(spec["substitutions"][init_time_spec_key])
        df = Psout(psout_path).to_df(secs_to_remove=init_time_sec)
        if df_manipulation_fn is not None:
            df = df_manipulation_fn(df)

        disturbance_cmd = ParsedDslSignal.parse(spec[disturbance_command_spec_key]).signal
        fault_type_cmd = ParsedDslSignal.parse(spec[fault_type_command_spec_key]).signal
        assert disturbance_cmd is not None
        assert fault_type_cmd is not None

        init_fault_times = [piece.time_start for piece in disturbance_cmd.pieces if piece.target==1.0]
        clear_fault_times = [piece.time_start for piece in disturbance_cmd.pieces if piece.target==0.0]
        init_clear_pairs = list(zip(init_fault_times,clear_fault_times))
        fault_types = [piece.target for piece in fault_type_cmd.pieces if piece.target != 0.0]

        # Convert these columns to an inverter trip flag once we know what the correct enums to use are
        err_msgs = df[err_msg_spec_key].unique() if err_msg_spec_key is not None else []
        op_states = df[operation_state_spec_key].unique() if err_msg_spec_key is not None else []


        def get_minimum_time_between_faults(clear_fault_times):
            time_between = []
            for index, item in enumerate(clear_fault_times):
                if index != len(clear_fault_times)-1:
                    time_between.append(round(clear_fault_times[index+1]-item,2))
            
            return min(time_between)
        
        def get_list_of_fault_types(fault_type_cmd):

            list_of_fault_types = [piece.target for piece in fault_type_cmd.pieces if piece.target!=0]
            enums = {
                1: "1PhG",
                4: "2PhG",
                7: "3PHG",
                8: "PhPh",
                11: "3Ph"
                }
            
            return ",".join([enums[item] for item in list_of_fault_types if item in enums])


        def get_number_faults_of_type(fault_type_cmd, enum_to_count):

            list_of_faults = [piece.target for piece in fault_type_cmd.pieces]

            return list_of_faults.count(enum_to_count)
            
        def get_disturbances_with_residual_below(df,spec_key,threshold):

            count_below_threshold = 0
            for (fault_start_time, fault_clear_time) in init_clear_pairs:
                df_snip = df[(df.index > fault_start_time) & (df.index < fault_clear_time)] 
                if len(df_snip[df_snip[spec_key] < threshold])!=0:
                    count_below_threshold = count_below_threshold + 1
            return count_below_threshold

        def get_disturbances_with_residual_between(df,spec_key,lower_threshold,upper_threshold):

            count_between_threshold = 0
            for (fault_start_time, fault_clear_time) in init_clear_pairs:
                df_snip = df[(df.index > fault_start_time) & (df.index < fault_clear_time)] 
                if len(df_snip[(df_snip[spec_key] > lower_threshold)&(df_snip[spec_key] < upper_threshold)])!=0:
                    count_between_threshold = count_between_threshold + 1
            return count_between_threshold            

        def get_deltav_deltat_below(
                df, 
                spec_key,
                threshold
                ):

            above_limit = df[spec_key] > threshold
            detect_cross_threshold = above_limit != above_limit.shift()
            detect_cross_threshold_cum_sum = detect_cross_threshold.cumsum()

            clusters = df[df[spec_key] < threshold].groupby(detect_cross_threshold_cum_sum)

            approx_area = 0
            for key, group in clusters:
                if len(group[spec_key])!=0:

                    # delta_t = group.index[1:]-group.index[:-1]
                    # delta_t = np.array(delta_t)
                    # delta_v = delta_v[:-1]
                    delta_v = threshold-np.array(group[spec_key])
                    t = group.index
                    approx_area = approx_area + np.trapz(delta_v,t)

            return round(approx_area,2)

        def get_cumulative_time_below(
                df, 
                spec_key,threshold
                ):

            above_limit = df[spec_key] > threshold
            detect_cross_threshold = above_limit != above_limit.shift()
            detect_cross_threshold_cum_sum = detect_cross_threshold.cumsum()

            clusters = df[df[spec_key] < threshold].groupby(detect_cross_threshold_cum_sum)

            cumulative_time = 0
            for _, group in clusters:
                cumulative_time = cumulative_time + group.index[-1]-group.index[0]

            return round(cumulative_time,2)
    
        spec_key = "Vrms_poc_pu"
        str_of_fault_types = get_list_of_fault_types(fault_type_cmd)
        minimum_time_between_faults = get_minimum_time_between_faults(clear_fault_times)
        number_3phase_faults = get_number_faults_of_type(fault_type_cmd, 7)
        disturbances_with_residual_below = get_disturbances_with_residual_below(df,spec_key,0.9)
        disturbances_with_residual_between = get_disturbances_with_residual_between(df,spec_key,0.7,0.9)
        time_integral_pus = get_deltav_deltat_below(df,spec_key,0.9)
        time_below_threshold = get_cumulative_time_below(df,spec_key,0.9)


        scenario_basename = os.path.basename(psout_path).split(".psout")[0]
        new_row = {
                "Name": scenario_basename,
                "Fault Times (s)": ",".join(str(round(x,2)) for x in init_fault_times),
                "Fault Duration (s)": ",".join(str(round(x,2)) for x in np.subtract(clear_fault_times,init_fault_times)),
                "Fault Types": str_of_fault_types,
                "Minimum Time Between Faults": minimum_time_between_faults,
                "Number of 3 Phase Faults": number_3phase_faults,
                "Disturbances with Residual Voltage Below 90%": disturbances_with_residual_below,
                "Disturbances with Residual Voltage Between 70-90%": disturbances_with_residual_between,
                "Time Below Vpoc of 90%": time_below_threshold,
                "Time Integral (pu.s)": time_integral_pus,
                "Inverter Operation States":",".join([str(x) for x in op_states]),
                "Inverter Error Messages": ",".join([str(x) for x in err_msgs]),
            }

        results_df = pd.concat([results_df, pd.DataFrame([new_row])], ignore_index=True)
    results_df.to_csv(output_csv_path, index=False)


