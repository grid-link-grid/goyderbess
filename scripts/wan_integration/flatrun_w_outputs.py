import os,sys

MAIN_DIRECTORY = os.getcwd()
sys_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSPY39'  #or where else you find the psspy.pyc
sys.path.append(sys_path_PSSE)
os_path_PSSE=r' C:\Program Files (x86)\PTI\PSSE34\PSSBIN'  # or where else you find the psse.exe
os.environ['PATH'] += ';' + os.environ['PATH']

import psse34
import psspy
from psspy import _i,_f,_s
from datetime import datetime
from gridlink.utils.wan.wan_utils import merge_jsons_at_paths,merge_dictionaries,find_files,exportToExcelSheets
import json


sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from flatrun_wan import flat_run
from WANPlotter import WANPlotter
from pallet.psse.Out import Out


results_dir = r"""C:\Grid\Heywood work folder\WAN\results"""
integrated_case_dir = r"""C:\Grid\WorkFolder\Heywood work folder\WAN_HighestRetune\HighestCase"""
CASE_NAME = "20241026-110045-subTrans-systemNormal_PEC3"
RUN_TIME_SECS = 2.0

TIME = datetime.now().strftime('%Y-%m-%d_%H-%M-%S-%f')[:-3]
    
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


def plot_outputs(plotdef_sub_dir: str,
         outfile_path: str,
         results_sub_dir: str):

    plot_def_paths = find_files(plotdef_sub_dir,".plotdef")
    #print(plot_def_paths)

    filename = os.path.basename(outfile_path).replace(".out","")
    df_from_out = Out(outfile_path).to_df()

    for i,path in enumerate(plot_def_paths):

        plotdef_name = os.path.basename(path).replace(".plotdef","")
        print(f"Processing plotdef: {plotdef_name}")
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

# merged_chandef_path = os.path.join(integrated_case_dir, "Model","wan.chandef")    
# with open(merged_chandef_path, 'w') as file:
#     json.dump(staged_chandef, file, indent=4) 

output_dir = os.path.join(results_dir,TIME + "_Dynamics")
dstate_df, ERR = flat_run(
    snapshot_dll_dir = r""".\\dsusr.dll""",
    case_dir = os.path.join(integrated_case_dir,"Model",CASE_NAME),
    dlls_dir = os.path.join(integrated_case_dir,"Model","DLLs"),
    dyrs_dir = os.path.join(integrated_case_dir,"Model","DYRs"),
    output_dir = output_dir,
    RUN_TIME = RUN_TIME_SECS
)

dstate_df.to_csv(os.path.join(integrated_case_dir,"dstate_errors.csv"), index=False)

#plot_outputs

# Create generic plotdef for generators and WAN
# merged_chandef_path = os.path.join(integrated_case_dir,"Model","wan.chandef")
# plotdef_subdir = os.path.join(integrated_case_dir,"PLOTDEFs")
# chandef_to_plotdef(merged_chandef_path,plotdef_subdir)

#Copy in any custom defs -- add later

# Plot outputs
plotdef_subdir = os.path.join(integrated_case_dir,"PLOTDEFs")
for outfile in find_files(output_dir,".out"):
    
    plot_outputs(plotdef_subdir, outfile, os.path.dirname(outfile))
