import sys

PF_PATH = r"C:\Program Files\DIgSILENT\PowerFactory 2024 SP4\Python\3.9"
sys.path.append(PF_PATH)
import powerfactory as pf

import os
import warnings
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.spatial import ConvexHull

warnings.simplefilter(action='ignore', category=pd.errors.PerformanceWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)

RESULTS_FOLDER = r"""C:\Grid\Results\Haywood"""
os.makedirs(RESULTS_FOLDER, exist_ok=True)

INPUTS_FOLDER = os.getcwd()
P_STEP_SIZE_MW = 0.01



def interpolate_q_cartesian(p_value, input_df):
    df = input_df.copy()
    df = df.sort_values(by="P_MW", ascending=True)
    df.reset_index(drop=True, inplace=True)

    print("--------------")
    # print(f"Pmax = {max(df['P_MW'])}, Pmin = {min(df['P_MW'])}")
    if p_value < min(df["P_MW"]):
        selected_p = min(df["P_MW"])
    elif p_value > max(df["P_MW"]):
        selected_p = max(df["P_MW"])
    else:
        selected_p = 999.9
    # print(f"Original = {p_value}, limited = {selected_p}")


    # Find the index of the nearest values to p_value
    idx = df['P_MW'].searchsorted(p_value)
    
    # If p_value is beyond the range of the DataFrame
    if idx == 0:
        q_value = df.at[0, 'Q_MVAr']
        print(f"BRANCH A: P = {p_value}\t\tQ = {q_value}")
        return q_value
    elif idx == len(df):
        q_value = df.at[len(df) - 1, 'Q_MVAr']
        print(f"BRANCH B: P = {p_value}\t\tQ = {q_value}")
        return q_value
    
    # Linearly interpolate the Q_MVAr value
    p0, p1 = df.at[idx - 1, 'P_MW'], df.at[idx, 'P_MW']
    q0, q1 = df.at[idx - 1, 'Q_MVAr'], df.at[idx, 'Q_MVAr']
    
    
    q_value = q0 + (q1 - q0) * (p_value - p0) / (p1 - p0)
    print(f"{p0=}, {p1=}")
    print(f"{q0=}, {q1=}")
    print(f"{q_value=}")
    print(f"P = {p_value}\t\tQ = {q_value}")
    
    return q_value





def generate_float_range(start, stop, step):
    # Calculate the number of steps needed to reach or slightly exceed 'stop'
    num_steps = int(np.ceil((stop - start) / step))
    
    # Generate the range of floats
    float_range = start + np.arange(num_steps) * step
    
    # Include 'stop' if it's not already included
    if float_range[-1] != stop:
        float_range = np.append(float_range, stop)
    
    return float_range


def generate_inverter_pq_curve_in_pf_format() -> pd.DataFrame:
    curves_df = pd.read_csv(os.path.join(INPUTS_FOLDER, "PQ_1725_RAW.csv"))
    p_range = generate_float_range(start=min(curves_df["P_MW"]), stop=max(curves_df["P_MW"]), step=P_STEP_SIZE_MW)


    data = {}
    p_mw_index = []

    for degc in [30, 45, 50]:
        for voltage in [0.85, 0.9, 1.0, 1.1]:
            filtered_curves_df = curves_df[(curves_df["degC"] == degc) & (curves_df["V_pu"] == voltage)]
            print(f"{degc}C, {voltage}pu: {len(filtered_curves_df)} entries")
            
            cap_curves_df = filtered_curves_df[filtered_curves_df["Q_MVAr"] >= 0]
            ind_curves_df = filtered_curves_df[filtered_curves_df["Q_MVAr"] <= 0]

            for hemisphere in ["Capacitive", "Inductive"]:

                pf_curve_p_mw =  []
                pf_curve_q_mvar = []
                
                for p_mw in p_range:
                    pf_curve_p_mw.append(p_mw)
                    if hemisphere == "Capacitive":
                        pf_curve_q_mvar.append(interpolate_q_cartesian(p_mw, cap_curves_df))
                    else:
                        pf_curve_q_mvar.append(interpolate_q_cartesian(p_mw, ind_curves_df))

                p_mw_index = pf_curve_p_mw

                data[(degc, voltage, hemisphere, 'Q_MVAr')] = pf_curve_q_mvar

                # print(list(zip(pf_curve_p_mw, pf_curve_q_mvar)))
                # plt.scatter(pf_curve_q_mvar, pf_curve_p_mw)

            both_hemispheres_df =pd.concat([cap_curves_df, ind_curves_df])
            # plt.scatter(cap_curves_df["Q_MVAr"], cap_curves_df["P_MW"])
            # plt.scatter(ind_curves_df["Q_MVAr"], ind_curves_df["P_MW"])
            # plt.scatter(both_hemispheres_df["Q_MVAr"], both_hemispheres_df["P_MW"], label=f"{voltage} pu")
            # plt.scatter(pf_curve_q_mvar, pf_curve_p_mw)

        # plt.xlabel('Q (MVAr)')
        # plt.ylabel('P (MW)')
        # plt.title(f'Power Curves at {degc}C')
        # plt.legend()
        # plt.savefig(os.path.join(RESULTS_FOLDER, f'interpolated_{degc}C.png'))
        # plt.clf()

    df = pd.DataFrame(data)

    # Clean up for presentation
    df.index = p_mw_index


    return df


def generate_visual_check_plots(df : pd.DataFrame):
    raw_curves_df = pd.read_csv(os.path.join(INPUTS_FOLDER, "PQ_1725_RAW.csv"))

    for degc in [30, 45, 50]:
        for voltage in [0.85, 0.9, 1.0, 1.1]:
            plt.figure(figsize=(10,6))
            plt.clf()

            # filtered_df = df[(df['degc'] == degc) & (df['voltage'] == voltage)]
            filtered_series = pd.concat([
                df[(degc, voltage, 'Capacitive', 'Q_MVAr')],
                df[(degc, voltage, 'Inductive', 'Q_MVAr')],
                ])

            filtered_raw_curves_df = raw_curves_df[(raw_curves_df["degC"] == degc) & (raw_curves_df["V_pu"] == voltage)]

            plt.scatter(filtered_series.values, filtered_series.index, label='Interpolated')
            plt.scatter(filtered_raw_curves_df['Q_MVAr'], filtered_raw_curves_df['P_MW'], label='raw', color='black')

            title = f'Check plot, {degc} degC and {voltage } pu'
            plt.title(title)
            plt.xlabel('Q (MVAr)')
            plt.ylabel('P (MW)')
            plt.legend()
            plt.savefig(os.path.join(RESULTS_FOLDER, f'{title}.png'))



if __name__ == "__main__":

    if not os.path.exists(RESULTS_FOLDER):
        os.makedirs(RESULTS_FOLDER)

    df = generate_inverter_pq_curve_in_pf_format()

    df_transposed = df.T
    df_transposed.to_csv(os.path.join(RESULTS_FOLDER, "powerfactory_inverter_pq_curves.csv"))

    generate_visual_check_plots(df)


