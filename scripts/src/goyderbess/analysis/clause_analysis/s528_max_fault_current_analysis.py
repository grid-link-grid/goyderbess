import os
import glob
import json
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from typing import List, Tuple, Optional, Callable
from tqdm.auto import tqdm
from pallet.pscad.Psout import Psout
from pallet.dsl.ParsedDslSignal import ParsedDslSignal
from pallet.FaultCode import FaultCode, parse_fault_code
from gridlink.plotting.characteristic_overlays import make_characteristic_overlay


def produce_s528_max_fault_current_outputs(
        psout_paths : List[str],
        output_png_path : str,
        output_csv_path : str,
        init_time_spec_key : str,
        ia_rms_ka_signal_name: str,
        ib_rms_ka_signal_name: str,
        ic_rms_ka_signal_name: str,
        x_label : str = "All fault studies",
        y_label : str = "Max current (kA)",
        title : str = "s5.2.8 Max Fault Current (excluding grid contribution)",
        df_manipulation_fn : Optional[Callable] = None,
        ):

    results_df = pd.DataFrame(columns=[
        "Name",
        "Max current (kA)",
        ])

    for i in tqdm(range(len(psout_paths)), desc="s5.2.8 Max Fault Current Studies"):
        psout_path = psout_paths[i]
        json_path = psout_path.split(".psout")[0] + ".json"

        with open(json_path, 'r') as f:
            spec = json.load(f)
        
        init_time_sec = float(spec["substitutions"][init_time_spec_key])

        df = Psout(psout_path).to_df(secs_to_remove=init_time_sec)
        if df_manipulation_fn is not None:
            df = df_manipulation_fn(df)

        # Find the fault type
        # Note that this will only work for studies where a constant fault type is
        # applied. For MFRT, this will need to be parsed as DSL
        fault_code = parse_fault_code(spec["Fault_Type_sig"])
        assert isinstance(fault_code, FaultCode)
        fault_type = fault_code.name


        new_row = {
                "Name": os.path.basename(psout_path).split(".psout")[0],
                "Fault Type": fault_type,
                "Max rms phase current (kA)": max(max(df[ia_rms_ka_signal_name]), max(df[ib_rms_ka_signal_name]), max(df[ic_rms_ka_signal_name])),
                "Max rms current phase A (kA)": max(df[ia_rms_ka_signal_name]),
                "Max rms current phase B (kA)": max(df[ib_rms_ka_signal_name]),
                "Max rms current phase C (kA)": max(df[ic_rms_ka_signal_name]),
                }
        results_df = pd.concat([results_df, pd.DataFrame([new_row])], ignore_index=True)


    results_df.to_csv(output_csv_path, index=False)

    a_data = [
            results_df[results_df["Fault Type"] == "LLLG"]["Max rms current phase A (kA)"],
            results_df[results_df["Fault Type"] == "LLG"]["Max rms current phase A (kA)"],
            results_df[results_df["Fault Type"] == "LG"]["Max rms current phase A (kA)"],
            results_df[results_df["Fault Type"] == "LL"]["Max rms current phase A (kA)"],
        ]

    b_data = [
            results_df[results_df["Fault Type"] == "LLLG"]["Max rms current phase B (kA)"],
            results_df[results_df["Fault Type"] == "LLG"]["Max rms current phase B (kA)"],
            results_df[results_df["Fault Type"] == "LG"]["Max rms current phase B (kA)"],
            results_df[results_df["Fault Type"] == "LL"]["Max rms current phase B (kA)"],
        ]

    c_data = [
            results_df[results_df["Fault Type"] == "LLLG"]["Max rms current phase C (kA)"],
            results_df[results_df["Fault Type"] == "LLG"]["Max rms current phase C (kA)"],
            results_df[results_df["Fault Type"] == "LG"]["Max rms current phase C (kA)"],
            results_df[results_df["Fault Type"] == "LL"]["Max rms current phase C (kA)"],
        ]

    ticks = [
            f'LLLG\n(n={len(results_df[results_df["Fault Type"] == "LLLG"])})',
            f'LLG\n(n={len(results_df[results_df["Fault Type"] == "LLG"])})',
            f'LG\n(n={len(results_df[results_df["Fault Type"] == "LG"])})',
            f'LL\n(n={len(results_df[results_df["Fault Type"] == "LL"])})',
            ]

    def set_box_color(bp, color):
        plt.setp(bp['boxes'], color=color)
        plt.setp(bp['whiskers'], color=color)
        plt.setp(bp['caps'], color=color)
        plt.setp(bp['medians'], color=color)

    fig, ax = plt.subplots(nrows=1, ncols=1, figsize=(8, 9), sharex=False)
    plt.title("Maximum current (kA)")

    bp_a = plt.boxplot(a_data, positions=np.array(range(len(a_data)))*2.0-0.4, sym='', widths=0.3) 
    bp_b = plt.boxplot(b_data, positions=np.array(range(len(a_data)))*2.0, sym='', widths=0.3) 
    bp_c = plt.boxplot(c_data, positions=np.array(range(len(a_data)))*2.0+0.4, sym='', widths=0.3) 
    set_box_color(bp_a, '#D7191C')
    set_box_color(bp_b, '#2C7BB6')
    set_box_color(bp_c, 'green')

    # Draw temporary lines
    plt.plot([], c='#D7191C', label='Phase A')
    plt.plot([], c='#2C7BB6', label='Phase B')
    plt.plot([], c='green', label='Phase C')
    plt.legend()

    plt.xticks(range(0, len(ticks) * 2, 2), ticks)
    plt.xlim(-2, len(ticks)*2)
    # plt.ylim()
    plt.tight_layout()

    ax.set_ylabel("Current (kA)")

    plt.title(title)
    plt.ylabel(y_label)
    plt.savefig(output_png_path)




