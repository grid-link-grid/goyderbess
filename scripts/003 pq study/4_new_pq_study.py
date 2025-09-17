import sys

PF_PATH = r"C:\Program Files\DIgSILENT\PowerFactory 2024 SP4\Python\3.9"
sys.path.append(PF_PATH)
import powerfactory as pf

import os
import warnings
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from typing import Dict, Optional
import statistics
from tqdm import tqdm
from cycler import cycler
from icecream import ic

from pallet.powerfactory.PowerfactoryNativePQPallet import PowerfactoryNativePQPallet, PDispatchStrategy, QDispatchStrategy
from pallet.powerfactory import statics
from pallet.logging import PalletLogger
from pallet.utils.time_utils import now_str
from pallet.utils.numpy_utils import generate_float_range

warnings.simplefilter(action='ignore', category=pd.errors.PerformanceWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)

logger = PalletLogger(name="user_script")

RESULTS_FOLDER = os.path.join(r"""C:\Grid\Results\Heywood\results\PQ-curve-R0""", "pq_curve_results_R0", now_str())

GEN_UNIT_CAPABILITY_MAPPING = {
        35.0: "Capability Curve_35deg",
        40.0: "Capability Curve_40deg",
        50.0: "Capability Curve_50deg",
        
    }

PBASE_MW = 300
QBASE_MVAR = PBASE_MW * 0.395
OLTC_MIN_TAP =1
OLTC_MAX_TAP =27
OLTC_TAP_LIMT = 2

AAS = pd.DataFrame({
        "P_MW": [PBASE_MW, -1*PBASE_MW, -1*PBASE_MW, PBASE_MW, PBASE_MW],
        "Q_MVAr": [QBASE_MVAR, QBASE_MVAR, -QBASE_MVAR, -QBASE_MVAR, QBASE_MVAR],
        })

colour_cycle = [
    '#0072B2',  # Blue
    '#D55E00',  # Vermillion
    '#009E73',  # Green
    '#CC79A7',  # Reddish Purple
    '#F0E442',  # Yellow
    '#56B4E9',  # Sky Blue
    '#E69F00',  # Orange
    # '#000000',  # Black
    # '#999999',  # Gray
    '#882255',  # Burgundy
    '#44AA99',  # Teal
    '#117733'   # Dark Green
]

PLOT_CUO = True


# *****************

PROJECT_NAME = "HeywoodBESS"
INV_QUERY ="Inverter*.ElmGenstat"
#WTG_QUERY = "*WTG*.ElmGenstat"
POC_BUS_QUERY='275kV BUS.ElmTerm'
EXT_GRID_QUERY='External Grid.ElmXnet'
GRID_BRANCH_QUERY='Grid Branch.ElmLne'
#METER_BRANCH_QUERY='Meter Branch.ElmLne'
OLTC_TRANSFORMERS_QUERY='TX*.ElmTr3'
#STATCOM_TRANSFORMERS_QUERY ="STATCOM TX*.ElmTr2"
#STATCOMS_QUERY = "*STATCOM*.ElmGenstat"
FILTERS_QUERY = "*.ElmFilter"                #Add Filter when available
LOADINGS_CABLES_QUERY = "*.ElmLne"
LOADINGS_TRANSFORMERS_QUERY = "*.ElmTr2"

METER_POINTS_TOWARDS_SLACK=False

# ****************

def sort_dataframe(df):
    sorted_df = df.copy()
    avg_qpoc_mvar = sorted_df['meas_qpoc_mvar'].mean()
    mask_low = sorted_df['meas_qpoc_mvar'] < avg_qpoc_mvar 
    mask_high = sorted_df['meas_qpoc_mvar'] >= avg_qpoc_mvar
    df_low = sorted_df[mask_low].sort_values(by='meas_ppoc_mw', ascending=True)
    df_high = sorted_df[mask_high].sort_values(by='meas_ppoc_mw', ascending=False)
    sorted_df = pd.concat([df_low, df_high])
    return sorted_df


def run_outer_capability_study(output_dir : str):
    app = pf.GetApplicationExt()
    # app.Show()

    statics.load_project(app=app, project_name=PROJECT_NAME, default_opscenario="Reactive Capability")
    logger.info("Loaded project")

    ext_grids = statics.locate_assets(app, query=EXT_GRID_QUERY, num_expected=1)
    poc_buses = statics.locate_assets(app, query=POC_BUS_QUERY, num_expected=1)
    grid_branches = statics.locate_assets(app, query=GRID_BRANCH_QUERY, num_expected=1)
    oltc_txs = statics.locate_assets(app, query=OLTC_TRANSFORMERS_QUERY, num_expected=2)
    invs = statics.locate_assets(app, query=INV_QUERY, num_expected=92)
    #wtgs = statics.locate_assets(app, query=WTG_QUERY, num_expected=56)
    #statcoms = statics.locate_assets(app, query=STATCOMS_QUERY, num_expected=1)
    #statcom_txs = statics.locate_assets(app, query=STATCOM_TRANSFORMERS_QUERY, num_expected=1)
    filters = statics.locate_assets(app, query=FILTERS_QUERY)
    if not filters:
        print("No filters found")
        filters = []
    loadings_cables = statics.locate_assets(app, query=LOADINGS_CABLES_QUERY)
    loadings_transformers = statics.locate_assets(app, query=LOADINGS_TRANSFORMERS_QUERY)
    ldf = statics.get_ldf(app)
    # statcom_tx_type = statics.get_2wnd_type(app, statcom_txs[0].GetAttribute("typ_id"))

    logger.info(f"Found {len(ext_grids)} external grid") 
    logger.info(f"Found {len(poc_buses)} POC bus") 
    logger.info(f"Found {len(grid_branches)} grid branch") 
    logger.info(f"Found {len(oltc_txs)} OLTC-equipped transformers")
    logger.info(f"Found {len(invs)} Inverters") 
    #logger.info(f"Found {len(wtgs)} turbines") 
    #logger.info(f"Found {len(statcoms)} STATCOMs") 
    #logger.info(f"Found {len(statcom_txs)} STATCOM transformers") 
    #logger.info(f"Found {len(filters)} filter") 
    logger.info(f"Found {len(loadings_cables)} cables")
    # logger.info(f"Found {1 if statcom_txs is not None else 0} STATCOM transformer types")

    oltc_tx_names = [tx.GetAttribute("e:loc_name") for tx in oltc_txs]
    loadings_cable_names = [cable.GetAttribute("e:loc_name") for cable in loadings_cables]
    loadings_transformers_names = [tx.GetAttribute("e:loc_name") for tx in loadings_transformers]

    def log_results(
            p_mw: float,
            vpoc_pu : float,
            degC: float,
            #tot_statcom_mvar_size : float,
            tot_filter_mvar_size : float,
            vcuo_pu : Optional[float] = None,
            ) -> Dict:
        result = dict()
        result["vpoc_pu"] = vpoc_pu
        result["vcuo_pu"] = vcuo_pu
        result["degC"] = degC
        #result["STATCOMs_total_mvar"] = tot_statcom_mvar_size
        result["Filters_total_mvar"] = tot_filter_mvar_size
        result["BESS_Pref_MW"] = p_mw
        result["BESS_Vref_pu"] = invs[0].GetAttribute("usetp")
        #result["WTG_Pref_MW"] = p_mw
        #result["WTG_Vref_pu"] = wtgs[0].GetAttribute("usetp")
        result["meas_vpoc_pu"] = poc_buses[0].GetAttribute("m:u")
        result["meas_qpoc_mvar"] = grid_branches[0].GetAttribute("m:Q:bus2")
        result["meas_ppoc_mw"] = grid_branches[0].GetAttribute("m:P:bus2")
    
        for oltc_idx, tx in enumerate(oltc_txs):
            tx_class_name = tx.GetClassName()
            if tx_class_name == "ElmTr2":
                result[f"TX_{oltc_tx_names[oltc_idx]}_ntap"] = int(tx.GetAttribute("c:nntap"))
                

            elif tx_class_name == "ElmTr3":
                # Note: The code is configured to only work for transformers controlling the HV bus. Further changes required if this isn't the case.
                result[f"TX_{oltc_tx_names[oltc_idx]}_ntap"] = int(tx.GetAttribute("n3tap_h"))
                # current_taps = int(tx.Ntap(0))
                # tx.SetAttribute("n3tap_h",current_taps)
                # print(current_taps)
                
                
            else:
                raise ValueError("Only 2 and 3 winding transformers have been implemented")
    
            cable_loadings = [cable.GetAttribute("c:loading") for cable in loadings_cables]
            cable_currents_ka = [cable.GetAttribute("m:I:bus1") for cable in loadings_cables]
            for x in range(len(loadings_cables)):
                result[f"CBL_{loadings_cable_names[x]}_loading_perc"] = cable_loadings[x]
                result[f"CBL_{loadings_cable_names[x]}_current_kA"] = cable_currents_ka[x]

            tx_loading = [tx.GetAttribute("c:loading") for tx in loadings_transformers if tx.GetAttribute("outserv") != 1]
            for x in range(len(loadings_transformers)):
                if loadings_transformers[x].GetAttribute("outserv") == 1:
                    result[f"TX_{loadings_transformers_names[x]}_loading_perc"] = 0.0
                else:
                    result[f"TX_{loadings_transformers_names[x]}_loading_perc"] = loadings_transformers[x].GetAttribute("c:loading")
        
        return result


    # Preparation of the model
    statics.set_branches_infinite(grid_branches)
    statics.set_sgens_to_constv(invs)
    #statics.set_sgens_to_constv(wtgs)
    #statics.set_sgens_to_constq(statcoms)
    #statics.set_sgens_pref_mw(statcoms, 0.0)

    # Allowable voltage ranges
    continuous_v_pu_range = (0.9, 1.1)
    # post_step_allowable_taps = 0
    #post_step_allowable_taps = 4
    #post_step_v_pu_range = (0.9 - 0.0125*post_step_allowable_taps, 1.1 + 0.0125*post_step_allowable_taps)
    post_step_v_pu_range =continuous_v_pu_range
    
    # Set up ranges
    degC_values = [35.0, 40.0, 50.0]
    p_mw_range = generate_float_range(start=-4.6, stop=4.6, step=0.5)
    vpoc_pu_values = [0.9, 1.0, 1.0, 1.0, 1.1]
    vcuo_pu_values = [None, None, 0.9, 1.1, None]

    # NO TAP
    # statcom_mvar_sizes = [0.0, 145.0]
    # filter_mvar_sizes = [0.0, 85.0]

    # 4 TAP
    #statcom_mvar_sizes = [80.0]
    #filter_mvar_sizes = [15]

    filter_mvar_sizes = [15] if filters else [0.0]

    num_studies = len(degC_values) * len(p_mw_range) * 2 * (len(vpoc_pu_values) + len(vcuo_pu_values) - vcuo_pu_values.count(None)) * len(degC_values) * len(filter_mvar_sizes)

    pbar = tqdm(total=num_studies)

    results = []
    num_failures = 0
    ic(type(invs))
    ic(invs)
    for degC in degC_values:
        capability_curve = statics.get_capability_curve(app=app, rel_path=GEN_UNIT_CAPABILITY_MAPPING[degC])
        ic(degC)
        ic(GEN_UNIT_CAPABILITY_MAPPING[degC])
        ic(capability_curve)
    
      
        statics.set_sgens_capability_curves(sgens=invs, capability_curve=capability_curve)
        
        for inv in invs:
            inv.SetAttribute("pQlimType", capability_curve)
      
        
        for inv in invs:
            cap_curve_inv=inv.GetAttribute("pQlimType")
            ic(cap_curve_inv)
        
        print(degC)
        filter_mvar_sizes = [15] if filters else [0.0]   #add filter size here
        for filter_mvar_size in (filter_mvar_sizes):
            print(filter_mvar_size)
            statics.set_shunts_mvar(filters, filter_mvar_size)
            #statics.set_sgens_q_limits(statcoms, qmin_mvar=-1*statcom_mvar_size, qmax_mvar=statcom_mvar_size)
            #statics.set_2wnd_type_rating_mva(statcom_txs, statcom_mvar_size)

            # print(f"STATCOM intended size = {statcom_mvar_size}")
            # print(f"{statcoms[0].GetAttribute('cQ_min')=}")
            # print(f"{statcoms[0].GetAttribute('cQ_max')=}")
            # print(f"{statcom_txs[0].GetAttribute('outserv')=}")
            # print(f"{statcom_txs[0].GetAttribute('typ_id').GetAttribute('strn')=}")


            for q_direction in [-1, 1]:
                #statics.set_sgens_qref_mvar(statcoms, q_direction * statcom_mvar_size)
                statics.set_sgens_vscheds(invs, continuous_v_pu_range[0] if q_direction == -1 else continuous_v_pu_range[1])
                
                # Reverse on the +Q side to maintain order
                # selected_p_mw_range = p_mw_range if q_direction == -1 else reversed(p_mw_range)
                selected_p_mw_range = p_mw_range
                for p_mw in selected_p_mw_range:
                    statics.set_sgens_pref_mw(invs, p_mw)

                    for vpoc_pu, vcuo_pu in zip(vpoc_pu_values, vcuo_pu_values):
                        """ Non-CUO Test """
                        statics.set_ext_grids_vscheds(ext_grids, vpoc_pu)
                        statics.configure_loadflow(
                                ldf=ldf,
                                honour_p_limits=True,
                                honour_q_limits=True,
                                tap_changers_enabled=True,
                                enable_pf_controllers=False,
                                degC=degC,
                                )
                        ldf.Execute()

                        if vcuo_pu is None:
                            try:
                                result = log_results(
                                    p_mw=p_mw,
                                    vpoc_pu=vpoc_pu,
                                    vcuo_pu=None,
                                    degC=degC,
                                    #tot_statcom_mvar_size=statcom_mvar_size,
                                    tot_filter_mvar_size=filter_mvar_size,
                                )
                                results.append(result)
                            except Exception as e:
                                print(" Failed")
                                num_failures += 1
                            pbar.update()
                            continue

                        pbar.update()

                        """ CUO Tests """
                        # here is where to save the transformer oltc positions
                        
                        
                        def set_tap_limits(transformer):
                            current_tap = int(transformer.GetAttribute("n3tap_h"))
                            print(f" {current_tap=}")
                            
                            min_limit= OLTC_MIN_TAP #int(transformer.GetAttribute("c:optapmin"))
                            max_limit = OLTC_MAX_TAP #int(transformer.GetAttribute("c:optapmax"))
                            allowed_range=OLTC_TAP_LIMT
                            lower_limit = max(current_tap-allowed_range, min_limit)
                            upper_limit = min(current_tap+allowed_range, max_limit)
                            
                            # lower_limit = max(current_tap, min_limit)
                            # upper_limit = min(current_tap, max_limit)
                            print(f"{lower_limit=},{upper_limit=}")
                            return lower_limit, upper_limit
                        
                        for tx in oltc_txs:
                            limits = set_tap_limits(tx)
                            tx.SetAttribute("optapmin_h",limits[0])
                            tx.SetAttribute("optapmax_h",limits[1])
                        
                        statics.set_oltcs_to_current_position(oltc_txs) #HERE is where to modify the transformers
                        statics.set_ext_grids_vscheds(ext_grids, vcuo_pu)

                        # Allow a wider range to permit tapping to return to the continuous range
                        statics.set_sgens_vscheds(invs, post_step_v_pu_range[0] if q_direction == -1 else post_step_v_pu_range[1])

                        statics.configure_loadflow(
                                ldf=ldf,
                                honour_p_limits=True,
                                honour_q_limits=True,
                                tap_changers_enabled=True, #this needs to be changed
                                enable_pf_controllers=False,
                                degC=degC,
                                )
                        ldf.Execute()
                        try:
                            result = log_results(
                                p_mw=p_mw,
                                vpoc_pu=vpoc_pu,
                                vcuo_pu=vcuo_pu,
                                degC=degC,
                                #tot_statcom_mvar_size=statcom_mvar_size,
                                tot_filter_mvar_size=filter_mvar_size,
                            )
                            results.append(result)
                        except Exception as e:
                                print("Failed")
                                num_failures += 1
                        
                        # here is where to return the tap changer limits to the original value
                        for tx in oltc_txs:
                            tx.SetAttribute("optapmin_h",OLTC_MIN_TAP)
                            tx.SetAttribute("optapmax_h",OLTC_MAX_TAP)
                        pbar.update()

    results_df = pd.DataFrame(results)
    print(f"no.of failures {num_failures}")
    
    # Grouping column
    def apply_group_name(row):
        vpoc_pu = row['vpoc_pu']
        vcuo_pu = row['vcuo_pu']
        row['grouping'] = f'{vpoc_pu} pu' if vcuo_pu is None or np.isnan(vcuo_pu) else f'CUO {vpoc_pu} pu to {vcuo_pu} pu'
        return row

    results_df = results_df.apply(apply_group_name, axis=1)

    results_df.to_csv(os.path.join(RESULTS_FOLDER, "outer_df.csv"))
    
    # Remove obvious calculation errors
    results_df = results_df[abs(results_df["meas_qpoc_mvar"]) < 500.0]
    results_df = results_df[abs(results_df["meas_ppoc_mw"]) < 500.0]

    # All plots should have the same axes to enable comparisons when toggling between them
    plot_q_min = min(results_df['meas_qpoc_mvar']) * 1.1  
    plot_q_max = max(results_df['meas_qpoc_mvar']) * 1.1
    plot_p_min = min(min(results_df['meas_ppoc_mw']) * 1.1, 0.0)
    plot_p_max = max(results_df['meas_ppoc_mw']) * 1.1

    for degC in degC_values:
        for filter_mvar_size in(filter_mvar_sizes):
            total_filter_mvar = filter_mvar_size * len(filters) if filters else 0.0

            for plot_cuo in [True, False]:

                #figure_df = results_df[np.isclose(results_df["Filters_total_mvar"], total_filter_mvar) & np.isclose(results_df["STATCOMs_total_mvar"], statcom_mvar) & (results_df["degC"] == degC)]
                #figure_df = results_df[np.isclose(results_df["Filters_total_mvar"], total_filter_mvar) & (results_df["degC"] == degC)]
                if "Filters_total_mvar" in results_df.columns:
                    figure_df = results_df[np.isclose(results_df["Filters_total_mvar"], total_filter_mvar) & (results_df["degC"] == degC)]
                else:
                    figure_df = results_df[results_df["degC"] == degC]

                print(f"{total_filter_mvar=}")
                print(f"{degC=}")   
                #print(f"{statcom_mvar=}")
                print(f"Located {len(figure_df)} rows!")
                print(figure_df["grouping"].unique())

                plt.rc('axes', prop_cycle=cycler('color', colour_cycle))
                plt.clf()
                plt.close('all')
                plt.figure(figsize=(12, 10))
                print(f"Total Filter MVAR: {total_filter_mvar}")
                
                for grouping in figure_df["grouping"].unique():
                    filtered_df = figure_df[figure_df["grouping"] == grouping]

                    if not plot_cuo and "CUO" in grouping:
                        continue

                    sorted_df = sort_dataframe(filtered_df)
                    sorted_df = pd.concat([sorted_df, sorted_df.iloc[0:1]], ignore_index=True)
                    plt.plot(sorted_df["meas_qpoc_mvar"], sorted_df["meas_ppoc_mw"], label=grouping)

                plt.plot(AAS["Q_MVAr"], AAS["P_MW"], color="black", label="AAS", linestyle='--')

                #title = f"BRWF2 reactive capability curve - {degC}C, {total_filter_mvar} MVAr filters + {statcom_mvar} MVAr STATCOMs"
                #title = f"Heywood reactive capability curve - {degC}C, {total_filter_mvar} MVAr filters"
                # plt.title(title)

                title = f"Haywood reactive capability curve - {degC}C"
                if filters:
                    title += f", {total_filter_mvar} MVAr filters"

                plt.xlabel("Reactive Power (MVAr)")
                plt.xlim(plot_q_min, plot_q_max)
                plt.ylim(plot_p_min, plot_p_max)
                plt.ylabel("Active Power (MW)")
                plt.grid(True)
                plt.legend(loc='center left', bbox_to_anchor=(1, 0.05))

                filename = title + ".png"
                plt.savefig(os.path.join(output_dir, f"{'CUO_' if plot_cuo else ''}{filename}"), bbox_inches='tight')




if __name__ == "__main__":
    if not os.path.exists(RESULTS_FOLDER):
        os.makedirs(RESULTS_FOLDER)

    run_outer_capability_study(output_dir=RESULTS_FOLDER)
    #print(f"{QBASE_MVAR=}")
    print("Heywood Reactive Capability curves generated")
    
    