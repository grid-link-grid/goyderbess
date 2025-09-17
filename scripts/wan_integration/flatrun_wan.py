import os
import sys
sys.path.insert(0,'C:\\Program Files (x86)\\PTI\\PSSE34\\PSSPY39')
import psse34
import psspy
from psspy import _i,_f,_s
import pandas as pd
import re
from parse_logfile import parse_logfile_dstates
import json
from WANPlotter import WANPlotter
from pallet.psse.Out import Out


def find_files(directory,file_type):
    file_dirs = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(file_type):
                file_dirs.append(os.path.join(root, file))
    return file_dirs


def combine_text_files(directories, output_file):
    with open(output_file, 'w') as outfile:
        for directory in directories:
            with open(directory, 'r') as infile:
                outfile.write("\n")
                outfile.write(directory)
                outfile.write("\n")
                outfile.write(infile.read())
                outfile.write("\n")


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


def add_channels(chandef_subdir):

    chandefs = find_files(chandef_subdir,".chandef")

    if len(chandefs)!=0:
        for chandef in chandefs:
            # Read the JSON file and load it as a dictionary
            with open(chandef, 'r') as file:
                dict = json.load(file)
            
            if "bus_voltage_channels" in dict:
                for item in dict["bus_voltage_channels"]:
                    psspy.voltage_and_angle_channel([-1,-1,-1,item["bus"]],[item["vol_name"],item["ang_name"]])
            
            # if "var_channels" in dict:
            #     for item in dict["var_channels"]:
            #         if item["model"]["model_type"] == "PLANT":
            #             psspy.machine_array_channel([-1,])
                    
            #         if item["model"]["model_type"] == "OTHER_BUS":


            if "branch_pq_channels" in dict:
                for item in dict["branch_pq_channels"]:
                    psspy.branch_p_and_q_channel([-1,-1,-1,item["branch"]["from_bus"],item["branch"]["to_bus"]],item["branch"]["circuit_id"],[item["p_name"],item["q_name"]])
    else:
        psspy.voltage_and_angle_channel([-1,-1,-1,369080],["SHTS_V","SHTS_ANG"])



    # psspy.voltage_and_angle_channel([-1,-1,-1,100],["GGS_POC_V","GGS_POC_ANG"])

    # psspy.bus_frequency_channel([-1,369080], "SHTS 220kV")   

    # psspy.branch_p_and_q_channel([-1,-1,-1,200,100],r"""1""",[r"""GESF_POC_P""",r"""GESF_POC_Q"""])
    # psspy.branch_p_and_q_channel([-1,-1,-1,601,501],r"""1""",[r"""GESF_INV1_P""",r"""GESF_INV1_Q"""])
    # psspy.branch_p_and_q_channel([-1,-1,-1,602,502],r"""1""",[r"""GESF_INV2_P""",r"""GESF_INV2_Q"""])

    # #BRANCHES - TRANSMISSION
    # psspy.branch_mva_channel([-1,-1,-1,int(369080),int(100)],"1","SHTS-GESF")
    # psspy.branch_p_and_q_channel([-1,-1,-1,369080,100],r"""1""",[r"""SHTS-GESF_P""",r"""SHTS-GESF_Q"""])

    # psspy.branch_mva_channel([-1,-1,-1,int(100),int(323080)],"1","GESF-DDTS")
    # psspy.branch_p_and_q_channel([-1,-1,-1,100,323080],r"""1""",[r"""GESF-DDTS_P""",r"""GESF-DDTS_Q"""])

    # psspy.branch_mva_channel([-1,-1,-1,int(369080),int(329080)],"1","SHTS-GNTS")
    # psspy.branch_p_and_q_channel([-1,-1,-1,369080,329080],r"""1""",[r"""SHTS-GNTS_P""",r"""SHTS-GNTS_Q"""])

    # psspy.branch_mva_channel([-1,-1,-1,int(369080),int(327080)],"1","SHTS-FVTS")
    # psspy.branch_p_and_q_channel([-1,-1,-1,369080,327080],r"""1""",[r"""SHTS-FVTS_P""",r"""SHTS-FVTS_Q"""])

    # psspy.branch_mva_channel([-1,-1,-1,int(327080),int(311080)],"1","FVTS-BETS")
    # psspy.branch_p_and_q_channel([-1,-1,-1,327080,311080],r"""1""",[r"""FVTS-BETS_P""",r"""FVTS-BETS_Q"""])




def flat_run(
        snapshot_dll_dir: str,
        case_dir : str,
        dlls_dir : str,
        dyrs_dir : str,
        output_dir: str,
        RUN_TIME: float
        ):
    
    print(f"{snapshot_dll_dir=}")
    print(f"{case_dir=}")
    print(f"{dlls_dir=}")
    print(f"{output_dir=}")
    print(f"{output_dir=}")

    FLAT_RUN_COMPLETE = False

    psspy.psseinit(50000)

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    log_file = os.path.join(output_dir,"simulation_progress_output.txt")

    # if not os.path.exists(log_file):
    #     os.makedirs(log_file)

    with open(log_file, "w") as file:
        # Write some text to the file
        file.write("PROGRESS OUTPUT\n")
        file.close()


    # Redirect output to the file
    ierr = psspy.report_output(2, log_file, [0, 0])   # Redirects reports
    ierr = psspy.progress_output(2, log_file, [0, 0])  # Redirects progress messages
    ierr = psspy.alert_output(2, log_file, [0, 0])    # Redirects alerts
    ierr = psspy.prompt_output(2, log_file, [0, 0])   # Redirects prompts 
    print(f"{ierr=}")
    pass
    psspy.case(case_dir)
   # psspy.set_model_debug_output_flag(1)


    #psspy.progress_output(2,os.path.join(output_dir,"output.pdv"),[_i,0])

    
    psspy.short_circuit_units(1)
    psspy.short_circuit_z_units(1)
    psspy.short_circuit_coordinates(1)
    psspy.short_circuit_z_coordinates(0)
    psspy.set_netfrq(1)
    psspy.cong(0)


    # Define the Tasmanian subsystem, initialise for load conversion, convert loads, post-processing housekeeping
    psspy.bsys(sid=0,numarea=1,areas=[7])
    psspy.conl(apiopt=1)
    psspy.conl(0,0,2,[0,0],[91.27, 19.36,-126.88, 188.43])

    # Define a subsystem for APD loads and convert loads
    psspy.bsys(sid=1,numbus=2,buses=[301020,301037])
    psspy.conl(1,0,2,[0,0],[52.75, 58.13, 5.97, 95.52])

    # Define a subsystem for Tomago loads and convert loads
    psspy.bsys(sid=1,numbus=1,buses=[277690])
    psspy.conl(1,0,2,[0,0],[86.63, 25.19, -378.97, 347.97])

    # Define a subsystem for Boyne Island loads and convert loads
    psspy.bsys(sid=1,numbus=2,buses=[440840,440841])
    psspy.conl(1,0,2,[0,0],[51.36, 59.32,-228.04, 254.01])

    # Convert remaining loads in NEM
    psspy.conl(0,1,2,[0,0],[ 100.0,0.0,-306.02, 303.0])
    psspy.conl(apiopt=3)


    psspy.ordr(0)
    psspy.fact()
    psspy.solution_parameters_4([_i,_i,600,_i,_i],[_f,_f,_f,_f,_f,_f, 0.2,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    psspy.dynamics_solution_param_2([1000,_i,_i,_i,_i,_i,_i,_i],[ 0.1, 0.0005, 0.001, 0.016, 0.001, 0.001, 0.1, 0.0005])

    for i in range(16):
        psspy.tysl(0)


    psspy.set_chnfil_type(0)
    # psspy.report_output(1,"",[0,0])
    # psspy.progress_output(1,"",[0,0])
    # psspy.alert_output(4,"",[0,0])
    # psspy.prompt_output(4,"",[0,0])


    # Add dlls
    psspy.dropmodellibrary(snapshot_dll_dir)
    ierr = psspy.addmodellibrary(snapshot_dll_dir)
    print(f'DLL: {snapshot_dll_dir}" \nSTATUS: {str(ierr)}') 

    files = find_files(dlls_dir,".dll")
    files_filtered = [file for file in files if "dsusr.dll"!=os.path.basename(file)]

    for i,file in enumerate(files_filtered):
        psspy.dropmodellibrary(file)
        ierr = psspy.addmodellibrary(file)
        print(f'DLL: {file}\nSTATUS: {str(ierr)}') 

    # Add dyrs


    files = find_files(dyrs_dir,".dyr")

    for i, file in enumerate(files):
        if i==0:
            ierr = psspy.dyre_new([_i,_i,_i,_i],file,"","","")
            print(f'DYR: {file}\nSTATUS: {str(ierr)}') 
        else:
            ierr = psspy.dyre_add([_i,_i,_i,_i],file,"","")
            print(f'DYR: {file}\nSTATUS: {str(ierr)}') 

    # dyda_ierr = psspy.dyda(_i, _i, [_i, _i, _i], 0, "inmem")

    psspy.delete_all_plot_channels()

    # Add Channels HERE #
    add_channels(os.path.dirname(case_dir))
    try:
        print("Attempting strt_2")

        ierr = psspy.strt_2([0,0],os.path.join(output_dir,"output.out"))
        # ierr = psspy.strt_2([0,0],os.path.join(output_dir,"output.out"))
        # ierr = psspy.strt_2([0,0],os.path.join(output_dir,"output.out"))
        # ierr = psspy.strt_2([0,0],os.path.join(output_dir,"output.out"))
        print(f"str2 ierr = {ierr}")
        psspy.report_output(1, "", [0, 0])
        psspy.progress_output(1, "", [0, 0])
        psspy.alert_output(1, "", [0, 0])
        psspy.prompt_output(1, "", [0, 0])
        flat_run_ierr = psspy.run(0, RUN_TIME,1000,1,1)
        print(f"flat_run_ierr = {flat_run_ierr}")
        ierr = psspy.close_powerflow()
        print("Powerflow closed with message: "+str(ierr))
        ierr = psspy.pssehalt_2()


        # CHECK FOR DSTATE ERRORS
        df = parse_logfile_dstates(log_file)
        


    except Exception as e:
        print(f"{e=}")
        flat_run_ierr = 1
        df = None


    return df,flat_run_ierr





