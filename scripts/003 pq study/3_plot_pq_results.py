
import os
import glob
import sys
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import seaborn as sns
from pallet.utils.time_utils import now_str
from cycler import cycler

INPUTS_DIR = r"""C:\Users\LukeHyett\Grid-Link\Projects - Documents\Atmos\Heywood BESS R0\02_Deliverables\009 PQ Study\PowerFactory_PQstudy\285MW1.06pu_WithFilter_150Mvar"""
OUTPUTS_DIR = r"""C:\Users\LukeHyett\Grid\Results\Heywood"""
LOCAL_DIR = os.getcwd()




PBASE_MW = 285.0
QBASE_MVAR = PBASE_MW * 0.395
#VNORM_PU = 1.015
#VNORM_LESS_10PC_PU = 0.9015
NUM_INVERTERS = 25
#P_LIMIT_TOLERANCE_MW = 0.02
UPPER_MVAR_LIMIT = 500
LOWER_MVAR_LIMIT = -500



AAS = pd.DataFrame({
        "P_MW": [PBASE_MW, -1*PBASE_MW, -1*PBASE_MW, PBASE_MW, PBASE_MW],
        "Q_MVAr": [QBASE_MVAR, QBASE_MVAR, -QBASE_MVAR, -QBASE_MVAR, QBASE_MVAR],
        })


color_cycle = [    '#0072B2',  # Blue
    '#D55E00',  # Vermillion
    '#009E73',  # Green
    '#CC79A7',  # Reddish Purple
    '#F0E442',  # Yellow
    '#56B4E9',  # Sky Blue
    '#E69F00',  # Orange
    '#000000',  # Black
    '#999999',  # Gray
    '#882255',  # Burgundy
    '#44AA99',  # Teal
    '#117733'   # Dark Green
]



def calc_inv_tx_loading_scaling(degC):
    """ Only accurate between 35C and 50C """
    overload = 1.3703703345 - 0.0074074067 * degC
    scal = 0.0 if overload == 0.0 else 1.0/overload
    return scal


def order_clockwise(
        df: pd.DataFrame,
        x_col_name: str,
        y_col_name: str,
        ) -> pd.DataFrame:
    use_origin = True
    if not use_origin:
        centroid = [df[x_col_name].mean(), df[y_col_name].mean()]
    else:
        centroid = [0,0]
    df['angle'] = np.arctan2(df[y_col_name] - centroid[1], df[x_col_name] - centroid[0])
    sorted_df = df.sort_values(by='angle')
    return sorted_df

def check_pmeas_validity(
        df : pd.DataFrame,
        vmeas_pu: float,
        pmeas_mw : float,
        **kwargs,
        ):
    """
    Filters the given DataFrame based on vmeas_pu, pmeas_mw, and any additional keyword arguments.

    Parameters:
    - df (pd.DataFrame): The DataFrame to filter.
    - vmeas_pu: Measured terminal voltage
    - pmeas_mw: Measured active power generated
    - kwargs: Additional column-value pairs to filter by.

    Returns:
    - pd.DataFrame: The filtered DataFrame.
    """
    print("---------------")
    filtered_df = df.loc[(df[list(kwargs)] == pd.Series(kwargs)).all(axis=1)]
    print(filtered_df)

    exact = df[df['vterm_pu'] == vmeas_pu]
    if len(exact) == 1:
        print(exact)
        valid = pmeas_mw <= exact.loc[0]['pmax_mw']
        print(f"valid = {valid}")
        return valid
    below = df[df['vterm_pu'] <= vmeas_pu].tail(1)
    above = df[df['vterm_pu'] >= vmeas_pu].head(1)

    print(below)
    print(type(below))

    print(below['vterm_pu'])
    v_dist_perc = (vmeas_pu - below['vterm_pu']) / (above['vterm_pu'] - below['vterm_pu'])
    print(f"{v_dist_perc=}")


#def filter_invalid_p_points(
#    df: pd.DataFrame,
#    capability_df : pd.DataFrame
#) -> pd.DataFrame:
    
   # def check_p_validity(row : pd.Series) -> bool:
        # Filter out points where the inverter it outside of its capability due to PF data entry limitations
        # print(row)
    #    vmeas_pu = row['INV_vterm_pu_avg']
    #    pmeas_mw = row['INV_pterm_mw_sum'] / NUM_INVERTERS
    #    qmeas_mvar = row['INV_qterm_mvar_sum'] / NUM_INVERTERS
    #    spec_degc = row['spec_degc']

    #    if qmeas_mvar >= 0:
    #        filt_cap_df = capability_df[(capability_df['degC'] == spec_degc) & (capability_df['Q_MVAr'] > 0)]
    #    else:
    #        filt_cap_df = capability_df[(capability_df['degC'] == spec_degc) & (capability_df['Q_MVAr'] < 0)]
    #    defined_voltages = sorted(filt_cap_df['V_pu'].unique())

    #    if len(defined_voltages) == 0:
    #        raise Exception("No voltages found in capability dataframe")

        # print(vmeas_pu)

        # Interpolate voltages
       #if vmeas_pu <= min(defined_voltages):
         #   v1 = min(defined_voltages)
         #   v2 = v1
         #   dist_voltage = 0.0
        #elif vmeas_pu >= max(defined_voltages):
         #   v1 = max(defined_voltages)
         #   v2 = v1
         #   dist_voltage = 0.0
        #else:
        #    v1 = max([v for v in defined_voltages if v <= vmeas_pu])
        #    v2 = min([v for v in defined_voltages if v > vmeas_pu])
        #    dist_voltage = (vmeas_pu - v1) / (v2 - v1)

        # print(f"{vmeas_pu=}, {v1=}, {v2=}, {dist_voltage=}")

        #v1_max_p_mw = max(filt_cap_df[filt_cap_df['V_pu'] == v1]['P_MW'])
        #v2_max_p_mw = max(filt_cap_df[filt_cap_df['V_pu'] == v2]['P_MW'])

        #v1_min_p_mw = min(filt_cap_df[filt_cap_df['V_pu'] == v1]['P_MW'])
        #v2_min_p_mw = min(filt_cap_df[filt_cap_df['V_pu'] == v2]['P_MW'])

        # print(f"{v1_max_p_mw=}")
        # print(f"{v2_max_p_mw=}")

        #interpolated_p_max_mw = dist_voltage * (v2_max_p_mw - v1_max_p_mw) + v1_max_p_mw
        #interpolated_p_min_mw = dist_voltage * (v2_min_p_mw - v1_min_p_mw) + v1_min_p_mw
        # print(f"{interpolated_p_max_mw}")

        #if pmeas_mw >= 0:
            #valid = pmeas_mw <= interpolated_p_max_mw + P_LIMIT_TOLERANCE_MW
            #Catch cases where the main transformer tap is on max boost and inverter terminal voltage is already high.
            #This filter has been devised empirically
            #valid = valid and not (row["TX_Grid Tx_ntap"]==1 and row["INV_vterm_pu_avg"]>1.0)
            # valid = valid and not (row["TX_Grid Tx2_ntap"]==1 and row["INV_vterm_pu_avg"]>1.0)

        #else:
      #     valid = pmeas_mw >= interpolated_p_min_mw - P_LIMIT_TOLERANCE_MW
      #     valid = valid and not (row["TX_Grid Tx_ntap"]==1 and row["INV_vterm_pu_avg"]>1.0)
            # valid = valid and not (row["TX_Grid Tx1_ntap"]==1 and row["INV_vterm_pu_avg"]>1.0)


     #   return valid

    #df['valid_p_output'] = df.apply(lambda row: check_p_validity(row), axis=1)
   # return df


def plot_pq(df : pd.DataFrame):
    df = df[(df['meas_qpoc_mvar'] <= UPPER_MVAR_LIMIT) & (df['meas_qpoc_mvar'] >= LOWER_MVAR_LIMIT)]
    q_min = min(df['meas_qpoc_mvar']) * 1.1
    q_max = max(df['meas_qpoc_mvar']) * 1.1
    p_min = min(df['meas_ppoc_mw']) * 1.1
    p_max = max(df['meas_ppoc_mw']) * 1.1


    # for degc in [50]:
    #for opscen in ["RC WITH_FILTER", "RC WITHOUT_FILTER"]:
    for degC in [35, 40, 50]:
            plt.rc('axes', prop_cycle=cycler('color', color_cycle))
            plt.clf()
            plt.close('all')
            plt.figure(figsize=(12, 10))
            for grouping in df['grouping'].unique():
                #filtered_df = df[(df["grouping"] == grouping) & (df['scenario'] == opscen) & (df['spec_degc'] == degc)]
                filtered_df = df[(df["grouping"] == grouping) & (df['degC'] == degC)]

                assert isinstance(filtered_df, pd.DataFrame)
                sorted_df = order_clockwise(filtered_df, "meas_qpoc_mvar", "meas_ppoc_mw")

                # Need to add the first row to the end or there will be a gap in the plot
                # if len(sorted_df) > 0:
                sorted_df = pd.concat([sorted_df, sorted_df.iloc[0:1]], ignore_index=True)

                # plt.scatter(sorted_df["meas_qpoc_mvar"], sorted_df["meas_ppoc_mw"], label=label, marker='')
                plt.plot(sorted_df["meas_qpoc_mvar"], sorted_df["meas_ppoc_mw"], label=grouping)

                
                # plt.scatter(filtered_df["meas_qpoc_mvar"],filtered_df["meas_ppoc_mw"])♠

            plt.plot(AAS["Q_MVAr"], AAS["P_MW"], color="black", label="AAS", linestyle='--')

            # Annotate corners of the AAS box
            corner_points = {
            f"({-QBASE_MVAR:.2f}, {PBASE_MW:.2f})": (-QBASE_MVAR, PBASE_MW),
            f"({QBASE_MVAR:.2f}, {PBASE_MW:.2f})": (QBASE_MVAR, PBASE_MW),
            f"({-QBASE_MVAR:.2f}, {-PBASE_MW:.2f})": (-QBASE_MVAR, -PBASE_MW),
            f"({QBASE_MVAR:.2f}, {-PBASE_MW:.2f})": (QBASE_MVAR, -PBASE_MW),
            }
            
            
            # for label, (x, y) in corner_points.items():
            #     plt.text(x, y, label, fontsize=10, ha='center', va='center')
            
            
            #title = f"LMTH reactive capability curve - {degc}C, {opscen.replace('_', ' ')}"
            title = f"Heywood reactive capability curve- {degC}C"
            plt.title(title)
            plt.xlabel("Reactive Power (MVAr)")
            plt.xlim(q_min, q_max)
            plt.ylim(p_min, p_max)
            plt.ylabel("Active Power (MW)")
            plt.grid(True)
            plt.legend(loc='center left', bbox_to_anchor=(1, 0.5))
            
            # plt.show()
            filename = title + ".png"
            if os.path.exists(os.path.join(OUTPUTS_DIR, filename)):
                os.remove(os.path.join(OUTPUTS_DIR, filename))
            plt.savefig(os.path.join(timestamped_output_dir, "new_"+filename), bbox_inches='tight')


def plot_cable_loadings(df : pd.DataFrame):
    sns.set(style='whitegrid')

    # Calculate the max values for the x axis limits
    df = df[(df['meas_qpoc_mvar'] <= UPPER_MVAR_LIMIT) & (df['meas_qpoc_mvar'] >= LOWER_MVAR_LIMIT)]
    limits_df = df.filter(regex=r"^CBL_.*_loading_perc$")
    max_loading = max(limits_df.max())
    
    for grouping in df['grouping'].unique():
        filtered_df = df[df["grouping"] == grouping]

        filtered_df = filtered_df.filter(regex=r"^CBL_.*_loading_perc$")
        filtered_df.columns = filtered_df.columns.str.replace(r'^CBL_(.*)_loading_perc$', r'\1', regex=True)
        filtered_df.columns = filtered_df.columns.str.replace('_', '-')
        #filtered_df = filtered_df.drop(columns=["Meter Branch", "Grid Branch"])

        # Sort columns by name
        filtered_df = filtered_df.reindex(sorted(filtered_df.columns), axis=1)

        # # Sort columns by max value
        # filtered_df = filtered_df[filtered_df.max().sort_values(ascending=False).index]

        # Sort columns by average value
        # filtered_df = filtered_df[filtered_df.mean().sort_values(ascending=False).index]

        # for opscen in ["WITHOUT_FILTER", "WITH_FILTER"]:
        #     for degc in [35, 45, 50]:

        plt.clf()
        plt.close('all')
        plt.figure(figsize=(12, 16))
        sns.boxplot(data=filtered_df, orient='h')


        # title = f"GESF reactive capability study cable loadings - {voltage} pu, {opscen}, {degc} C"
        title = f"HYBESS reactive capability study cable loadings - {grouping}"
        filename = title + ".png"
        # plt.xticks(rotation=90)
        plt.gca().grid(True, which='both', axis='both')
        plt.xlabel("Loading (%)")
        plt.xlim(0, max_loading*1.1)
        plt.title(title)

        if os.path.exists(os.path.join(OUTPUTS_DIR, filename)):
            os.remove(os.path.join(OUTPUTS_DIR, filename))
        plt.savefig(os.path.join(timestamped_output_dir, filename), bbox_inches='tight')



def plot_tx_loadings(df : pd.DataFrame):
    sns.set(style='whitegrid')

    # Calculate the max values for the x axis limits
    df = df[(df['meas_qpoc_mvar'] <= UPPER_MVAR_LIMIT) & (df['meas_qpoc_mvar'] >= LOWER_MVAR_LIMIT)]
    limits_df = df.filter(regex=r"^TX_.*_loading_perc$")
    max_loading = max(limits_df.max())
    
    for grouping in df['grouping'].unique():
        filtered_df = df[df["grouping"] == grouping]

        filtered_df = filtered_df.filter(regex=r"^TX_.*_loading_perc$")

        filtered_df.columns = filtered_df.columns.str.replace(r'^TX_(.*)_loading_perc$', r'\1', regex=True)
        filtered_df.columns = filtered_df.columns.str.replace('_', '-')
        filtered_df = filtered_df.reindex(sorted(filtered_df.columns), axis=1)
        # filtered_df["plot-voltage"] = np.where(df['vcuo_pu'].isna(), df['spec_voltage'], df['vcuo_pu'])

        # Sort columns by name
        filtered_df = filtered_df.reindex(sorted(filtered_df.columns), axis=1)

        # # Sort columns by max value
        # filtered_df = filtered_df[filtered_df.max().sort_values(ascending=False).index]

        # Sort columns by average value
        # filtered_df = filtered_df[filtered_df.mean().sort_values(ascending=False).index]

        # voltage_df = filtered_df[(filtered_df["plot-voltage"] == voltage)]
        # assert isinstance(voltage_df, pd.DataFrame)
        # voltage_df = voltage_df.drop(columns=["plot-voltage"])
        plt.clf()
        plt.close('all')
        plt.figure(figsize=(12, 16))
        sns.boxplot(data=filtered_df, orient='h')
        plt.xlabel("Loading (%)")
        plt.xlim(0, max_loading*1.1)

        title = f"HYBESS reactive capability curve transformer loading - {grouping}"
        plt.title(title)

        filename = title + ".png"
        if os.path.exists(os.path.join(OUTPUTS_DIR, filename)):
            os.remove(os.path.join(OUTPUTS_DIR, filename))
        plt.savefig(os.path.join(timestamped_output_dir, filename), bbox_inches='tight')



def export_to_aemo_format_xlsx(df : pd.DataFrame):
    # df = input_df.set_index(['scenario', 'spec_voltage'])
    # df = df[['meas_ppoc_mw', 'meas_qpoc_mvar']]
    print(df.head())

    data = {}

    #for opscen in ["RC WITH_FILTER", "RC WITHOUT_FILTER"]:
        
    for degC in [35, 45, 50]:
            for grouping in df['grouping'].unique():
                #filtered_df = df[(df['scenario'] == opscen) & (df["spec_degc"] == degc) & (df["grouping"] == grouping)]
                filtered_df = df[(df["degC"] == degC) & (df["grouping"] == grouping)]

                #if len(filtered_df) > 0:
                #    data[(opscen, degc, grouping, 'P_MW')] = list(filtered_df['meas_ppoc_mw'].values)
                #    data[(opscen, degc, grouping, 'Q_MVAr')] = list(filtered_df['meas_qpoc_mvar'].values)
                
                if len(filtered_df) > 0:
                    data[(degC, grouping, 'P_MW')] = list(filtered_df['meas_ppoc_mw'].values)
                    data[(degC, grouping, 'Q_MVAr')] = list(filtered_df['meas_qpoc_mvar'].values)
    

    # Find the maximum length of the lists
    max_len = max(len(v) for v in data.values())
    print(f"{max_len=}")

    for k in data.keys():
        print(f"{k}: {len(data[k])} items")
        if len(data[k]) == 0:
            data[k] = [np.nan] * max_len
        elif len(data[k]) == max_len:
            pass
        else:
            print(type(data[k]))
            data[k] = data[k] + [None] * (max_len - len(data[k])) 


    for k in data.keys():
        print(f"{k}: {len(data[k])} items")

    # Create the DataFrame
    df = pd.DataFrame(data)

    df.to_csv(os.path.join(OUTPUTS_DIR, 'capability_curves_aemo_format.csv'))

    

def scale_tx_loadings(df : pd.DataFrame) -> pd.DataFrame:
    def scale_row(row):
        for col_name in df.columns:
            if col_name.startswith('TX') and col_name.endswith('_loading_perc'):
                degC = row['degC']
                row[col_name] = row[col_name] * calc_inv_tx_loading_scaling(degC)
        return row

    return df.apply(scale_row, axis=1)


def create_grouping_column(df : pd.DataFrame) -> pd.DataFrame:
    def apply_group_name(row):
        vpoc_pu = row['vpoc_pu']
        vcuo_pu = row['vcuo_pu']
        row['grouping'] = f'{vpoc_pu} pu' if vcuo_pu is None or np.isnan(vcuo_pu) else f'CUO {vpoc_pu} pu to {vcuo_pu} pu'
        return row

    return df.apply(apply_group_name, axis=1)

def get_most_recent_file(keywords):
    csv_files = glob.glob(os.path.join(INPUTS_DIR, "**", "*.csv"), recursive=True)
    
    #csv_files = glob.glob(os.path.join(INPUTS_DIR, "*.csv"))
    # Filter for files containing both keywords in the filename
    filtered_files = [f for f in csv_files if all(keyword in os.path.basename(f) for keyword in keywords)]
    print(filtered_files)
    most_recent_file = max(filtered_files, key=os.path.getctime)
    return most_recent_file

if __name__ == "__main__":
    if not os.path.exists(OUTPUTS_DIR):
        os.makedirs(OUTPUTS_DIR)

    timestamped_output_dir = os.path.join(OUTPUTS_DIR, f"{now_str()}")
    if not os.path.exists(timestamped_output_dir):
        os.makedirs(timestamped_output_dir)

    INNER_MODE = False

    if INNER_MODE:
        # Inner curve results file
        keywords = ['inner', 'FINAL']
        # most_recent_file = get_most_recent_file(keywords)
        most_recent_file = r"C:\Grid\cgbess\scripts\0012_updated_pq_curve_dec_2024\pq_curve_results_R1\20250505_1102_18149971\outer_df.csv"
        df = pd.read_csv(os.path.join(INPUTS_DIR, most_recent_file))
    else:
        keywords = ['outer', 'FINAL']
        # most_recent_file = get_most_recent_file(keywords)
        most_recent_file = r"outer_df_pointsremoved.csv"
        df = pd.read_csv(os.path.join(INPUTS_DIR, most_recent_file))
        # Outer curve results file
        # df = pd.read_csv(os.path.join(INPUTS_DIR, "_results_inner_20241003_1510_43911033_FINAL.csv"))

    #capability_df = pd.read_csv(os.path.join(LOCAL_DIR, "PQ_1725_RAW.csv"))

    #results_df = filter_invalid_p_points(df, capability_df=capability_df)
    results_df = create_grouping_column(df)
    results_df.to_csv(os.path.join(timestamped_output_dir, "capability_curve_results_inc_invalid.csv"))
    # results_df = results_df[results_df['valid_p_output'] == True]

    if INNER_MODE:
        # Stripping values below Pmax allows us to generate slighly more points
        # than required to ensure no corner cuts, then strip the excess
        results_df = results_df[results_df["meas_ppoc_mw"] <= PBASE_MW]

    # results_df = scale_tx_loadings(results_df)
    print(list(results_df.columns))

    results_df.to_csv(os.path.join(timestamped_output_dir, "capability_curve_results.csv"))

    export_to_aemo_format_xlsx(results_df)
    plot_pq(results_df)
    # plot_cable_loadings(results_df)
    # plot_tx_loadings(results_df)


