# take as an input the files (or their paths) that want to be analysed for model alignment
# parse these .pkl or .psout files for their last values (from PPOC, VPOC, QPOC, and meters)
# calculate losses using those extracted values
# output the data to a CSV in the same format as the PF model alignment output
# from cgbess.utils.extract_pscad_benchmark_values import 

import pandas as pd
import os
import glob
from pallet.pscad.Psout import Psout
from pallet import now_str



columns = [
    "Ppoc",
    "Qpoc",
    "Vpoc",
    "PcvtLV1",
    "QcvtLV1",
    "VcvtLV1",
    "PcvtLV2",
    "QcvtLV2",
    "VcvtLV2",
    "PcvtHV1",
    "QcvtHV1",
    "VcvtHV1",
    "PcvtHV2",
    "QcvtHV2",
    "VcvtHV2",
    "Ppi1",
    "Qpi1",
    "Vpi1",
    "Ppi2",
    "Qpi2",
    "Vpi2",
    "Phf1",
    "Qhf1",
    "Vhf1",
    "Phf2",
    "Qhf2",
    "Vhf2",
    "Phf3",
    "Qhf3",
    "Vhf3",
    "Phf4",
    "Qhf4",
    "Vhf4",
]

def final_values_to_csv(extension, inputs_dir, output_csv_dir):
    """put description here"""

    if extension != ".psout" and extension != ".pkl":
        print('ERROR: "extension" argument must be either ".pkl" or ".psout"')
        return

    results = pd.DataFrame([])

    input_paths = glob.glob(inputs_dir + "/*"+extension)

    for input_path in input_paths:
        if extension == ".psout":
            #read psout into dataframe
            input_df = Psout(input_path).to_df(secs_to_remove=0)
            
        if extension == ".pkl":
            #read pkl into dataframe:
            input_df = pd.read_pickle(input_path)
        
        results = pd.concat([results,input_df.iloc[len(input_df)-1]],ignore_index=True, axis=1)

    results = results.filter(columns, axis=0)
    results = results.T
    
    results["CONVERTER_TX_PLOSS_1"] = results["PcvtLV1"]-results["PcvtHV1"]
    results["CONVERTER_TX_QLOSS_1"] = results["QcvtLV1"]-results["QcvtHV1"]
    results["CONVERTER_TX_PLOSS_2"] = results["PcvtLV2"]-results["PcvtHV2"]
    results["CONVERTER_TX_QLOSS_2"] = results["QcvtLV2"]-results["QcvtHV2"]

    results["HALF_CABLE_PLOSS_1"] = results["PcvtHV1"]-results["Ppi1"]
    results["HALF_CABLE_QLOSS_1"] = results["QcvtHV1"]-results["Qpi1"]
    results["HALF_CABLE_PLOSS_2"] = results["PcvtHV2"]-results["Ppi2"]
    results["HALF_CABLE_QLOSS_2"] = results["QcvtHV2"]-results["Qpi2"]

    results["CABLE_PLOSS"] = results["HALF_CABLE_PLOSS_1"]+results["HALF_CABLE_PLOSS_2"]
    results["CABLE_QLOSS"] = results["HALF_CABLE_QLOSS_1"]+results["HALF_CABLE_QLOSS_2"]
    results["SHUNT_ABSORB_P"]=results["Phf1"]+results["Phf2"]+results["Phf3"]+results["Phf4"]
    results["SHUNT_ABSORB_Q"]=results["Qhf1"]+results["Qhf2"]+results["Qhf3"]+results["Qhf4"]
    results["GRID_TX_PLOSS_NO_SHUNT"]=(results["Ppi1"]+results["Ppi2"])-results["Ppoc"]
    results["GRID_TX_PLOSS_WITH_SHUNT"]=((results["Ppi1"]+results["Ppi2"])-results["SHUNT_ABSORB_P"])-results["Ppoc"]
    results["GRID_TX_QLOSS_NO_SHUNT"]=(results["Qpi1"]+results["Qpi2"])-results["Qpoc"]
    results["GRID_TX_QLOSS_WITH_SHUNT"]=((results["Qpi1"]+results["Qpi2"])-results["SHUNT_ABSORB_Q"])-results["Qpoc"]

    results.insert(value=(results.index+1),loc=0,column="Test Number")
    results.set_index("Test Number",drop=True,inplace=True)

    columns_to_drop = columns
    columns_to_drop.remove("Qpoc")
    columns_to_drop.remove("Ppoc")
    columns_to_drop.remove("Vpoc")
    results = results.drop(columns=columns_to_drop)
    results = results.drop(columns=["HALF_CABLE_PLOSS_1","HALF_CABLE_QLOSS_1","HALF_CABLE_PLOSS_2","HALF_CABLE_QLOSS_2"])

    NOW_STR = now_str()
    results.to_csv(os.path.join(output_csv_dir,f"pscad_results{NOW_STR}.csv"))
    


    return


#debugging section: remove once integrated into a larger script:

final_values_to_csv(".psout",r"C:\Grid\cg\pscad_results\_pscad_results_20250313_1135_16522563\Fgrid Steps",r"C:\Grid\cg\pscad_pf_model_alignment_results")