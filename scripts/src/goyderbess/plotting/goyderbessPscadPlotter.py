"""
This is the class for the new plotter for Learmonth BESS - DMAT studies PSCAD

"""



from typing import Optional, Union, Dict, List, Callable
from pathlib import Path
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.gridspec as gridspec
import pandas as pd
import numpy as np
from matplotlib.table import Cell, Table
import textwrap as twp
from dataclasses import dataclass
from pallet.plotting import Plotter
import warnings
import os
import sys
import ast
import re



class goyderbessPscadPlotter(Plotter): 
    def __init__(
            self,
            remove_first_seconds:bool

    ):
        self.remove_first_seconds = remove_first_seconds

   
    def plot_from_df_and_dict(
            self,
            df: pd.DataFrame,
            scenario_dict: Optional[Dict] = None,
            png_path: Optional[Union[str, Path]] = None,
            pdf_path: Optional[Union[str, Path]] = None,
    ):
       # print(df.columns.tolist())
        # print(df["Fslack_meas_Hz"])
        # print(df['Fslack_meas_Hz'].dtype)
        # print(df['Fslack_meas_Hz'].head())
        # print(df['Id_pos_pu'])
        # print(df['Id_pos_pu'].dtype)
        print(df.columns.to_list())
        #print(scenario_dict.get("substitutions"))

        plt.clf()
        plt.close()
        matplotlib.use('Agg')

        col_pallet = ["#ab427a",
                "#96b23d",
                "#6162b9",
                "#ce9336",
                "#bc6abc",
                "#5aa554",
                "#ba4758",
                "#46c19a",
                "#b95436",
                "#9e8f41"]
        
        if self.remove_first_seconds:
            REMOVE_FIRST_SECONDS = float(scenario_dict["substitutions"]["TIME_Full_Init_Time_sec"])
            df = df[df.index > REMOVE_FIRST_SECONDS]
            df.index = df.index - REMOVE_FIRST_SECONDS
        

# -------------------------------------------------SUBPLOT CREATION ------------------------------------------------------
        fig, axes = plt.subplots(5, 3, figsize=(10, 12),squeeze = False, sharex=True)
        cm = 1 / 2.54
        fig.set_size_inches(42 * cm, 29.7 * cm)  # A4 Size Landscape.
        fig.suptitle(scenario_dict["File_Name"],fontweight = 'bold')
        fig.subplots_adjust(left=0.05, right=0.95, bottom=0.05, top=0.9, wspace=0.2, hspace=0.3)
        axes[0,0].set_title("POC",loc = 'center',fontsize=13, pad=12)
        axes[0,1].set_title("INV", loc = 'center',fontsize=13, pad=12)
        axes[1,2].set_title('MISC',loc='center',fontsize=13, pad=12)
        axes[0][2].axis('off')

#------------------- ---------------------------------POC PLOTS -----------------------------------------------------------
        # [0, 0] POC active power
        axes[0,0].plot(df.index, df["Pdir_REF"], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[0,0].xaxis.set_label_position('top')
        axes[0,0].plot(df.index, df["P_bes_meter"], color=col_pallet[1], label='MEAS',linewidth=1.5)

        axes[0,0].set_xlabel("Active Power (MW)",loc="left")
        axes[0,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[0,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[0,0]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = min(df["Pdir_REF"].min(),df["P_bes_meter"].min())
        max_val = max(df["Pdir_REF"].max(),df["P_bes_meter"].max())
        ax.set_ylim(min_val-10,max_val+10)


        # [1,0] POC Reactive power
        # axes[1,0].plot(df.index, df["QRMTCMDOUT"], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[1,0].plot(df.index, df["Qdir_REF"], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[1,0].xaxis.set_label_position('top')
        axes[1,0].plot(df.index, df["Q_bes_meter"], color=col_pallet[1], label='MEAS',linewidth=1.5)

        axes[1,0].set_xlabel("Reactive Power (MVAr)",loc="left")
        axes[1,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[1,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[1,0]
        ax.set_facecolor((0.98,0.98,0.98))   
        # min_val = min(df["QRMTCMDOUT"].min(),df["PLANT_Q_HV"].min())
        # max_val = max(df["QRMTCMDOUT"].max(),df["PLANT_Q_HV"].max())
        min_val = min(df["Qdir_REF"].min(),df["Q_bes_meter"].min())
        max_val = max(df["Qdir_REF"].max(),df["Q_bes_meter"].max())
        ax.set_ylim(min_val-1,max_val+1)

        # [2,0] POC Voltage  
        # axes[2,0].plot(df.index, df["VCMDOUT"], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[2,0].plot(df.index, df["Vsite_REF_pu"], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[2,0].plot(df.index, df["V_bes_meter_rms"], color=col_pallet[1], label='MEAS',linewidth=1.5)
        # if "Unbalanced" in png_path:
        #     axes[2,0] .plot(df.index, df["V_hv_rms_phase_A"], color=col_pallet[1], label='A',linewidth=1.5)
        #     axes[2,0] .plot(df.index, df["V_hv_rms_phase_B"], color=col_pallet[2], label='B',linewidth=1.5)
        #     axes[2,0] .plot(df.index, df["V_hv_rms_phase_C"], color=col_pallet[3], label='C',linewidth=1.5)
        axes[2,0].xaxis.set_label_position('top')
        # axes[2,0].axhline(y=float(scenario_dict["substitutions"]["NOTE_PPC_HVRT_pu"]), color='black', label='FREEZE THRESHOLD',linestyle="dashdot",linewidth=1.5)
        # axes[2,0].axhline(y=float(scenario_dict["substitutions"]["NOTE_PPC_LVRT_pu"]), color='black',linestyle="dashdot",linewidth=1.5)

        axes[2,0].set_xlabel("Voltage (pu)",loc="left")
        axes[2,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[2,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[2,0]
        ax.set_facecolor((0.98,0.98,0.98))
        # if "Unbalanced" in png_path:
        #     min_val = min(
        #         df["Vref_pu"].min(),
        #         df["V_hv_rms_phase_A"].min(),
        #         df["V_hv_rms_phase_B"].min(),
        #         df["V_hv_rms_phase_C"].min()  # Assuming you meant phase C, not repeating phase B
        #     )
        #     max_val = max(
        #         df["Vref_pu"].max(),
        #         df["V_hv_rms_phase_A"].max(),
        #         df["V_hv_rms_phase_B"].max(),
        #         df["V_hv_rms_phase_C"].max()
        #     )
        # else:
        min_val = min(df["Vsite_REF_pu"].min(),df["V_bes_meter_rms"].min())
        max_val = max(df["Vsite_REF_pu"].max(),df["V_bes_meter_rms"].max())
        ax.set_ylim(min_val-0.012,max_val+0.012)


        # [3,0] POC Pos Sequence Active Current (pu) 
        axes[3,0].plot(df.index, df["Id_pos_poc"], color=col_pallet[0], label='Id',linewidth=1.5)
        axes[3,0].xaxis.set_label_position('top')

        axes[3,0].set_xlabel("Pos Sequence Active Current (pu)",loc="left")
        axes[3,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[3,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[3,0]
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = df["Id_pos_poc"].min()
        max_val = df["Id_pos_poc"].max()
        ax.set_ylim(min_val-0.1,max_val+0.1)



        # [4,0] POC Pos Sequence Active Current (pu) 
        axes[4,0].plot(df.index, df["Iq_pos_poc"], color=col_pallet[0], label='Iq',linewidth=1.5)
        axes[4,0].xaxis.set_label_position('top')

        axes[4,0].set_xlabel("Pos Sequence Reactive Current (pu)",loc="left")
        axes[4,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[4,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[4,0]
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = df["Iq_pos_poc"].min()
        max_val = df["Iq_pos_poc"].max()
        ax.set_ylim(min_val-0.1,max_val+0.1)

#-----------------------------------------INV PLOTS----------------------------------------------------------------
        # [0,1] INV active Power 
        scaling = [0.0000927,1.0]
        axes[0,1].plot(df.index, df["pcomd:2"]*scaling[0], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[0,1].xaxis.set_label_position('top')
        axes[0,1].plot(df.index, df["P_inv_meter"]*scaling[1], color=col_pallet[1], label='MEAS',linewidth=1.5)


        axes[0,1].set_xlabel("Active Power (MW)",loc="left")
        axes[0,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[0,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[0,1]
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = min(
            (df["pcomd:2"] * scaling[0]).min(),
            (df["P_inv_meter"] * scaling[1]).min()
        )
        max_val = max(
            (df["pcomd:2"] * scaling[0]).max(),
            (df["P_inv_meter"] * scaling[1]).max()
        )
        ax.set_ylim(min_val-1,max_val+1)



        # [1,1] INV reactive Power 
        scaling = [0.0001, 1.0]
        axes[1,1].plot(df.index, df["qcomd:2"]*scaling[0], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[1,1].xaxis.set_label_position('top')
        axes[1,1].plot(df.index, df["Q_inv_meter"]*scaling[1], color=col_pallet[1], label='MEAS',linewidth=1.5)


        axes[1,1].set_xlabel("Reactive Power (MVAr)",loc="left")
        axes[1,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[1,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[1,1]
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = min(
            (df["qcomd:2"] * scaling[0]).min(),
            (df["Q_inv_meter"] * scaling[1]).min()
        )
        max_val = max(
            (df["qcomd:2"] * scaling[0]).max(),
            (df["Q_inv_meter"] * scaling[1]).max()
        )
        # ax.set_ylim(min_val-20,max_val+20)

        # [2,1] INV voltage
        axes[2,1] .plot(df.index, df["V_inv_meter_rms"], color=col_pallet[0], label='Vrms',linewidth=1.5)
        # if "Unbalanced" in png_path:
        #     axes[2,1] .plot(df.index, df["PCU2_Vll_rms_LV_AB_pu"], color=col_pallet[1], label='AB',linewidth=1.5)
        #     axes[2,1] .plot(df.index, df["PCU2_Vll_rms_LV_BC_pu"], color=col_pallet[2], label='BC',linewidth=1.5)
        #     axes[2,1] .plot(df.index, df["PCU2_Vll_rms_LV_CA_pu"], color=col_pallet[3], label='CA',linewidth=1.5)
        
        axes[2,1] .xaxis.set_label_position('top')
        axes[2,1] .set_xlabel("Voltage (pu)",loc="left")
        axes[2,1] .legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[2,1] .grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[2,1] 
        ax.set_facecolor((0.98,0.98,0.98))

        # if "Unbalanced" in png_path:
        #     min_val = min(
        #         df["VMEAS_LCL"].min(),
        #         df["PCU2_Vll_rms_LV_AB_pu"].min(),
        #         df["PCU2_Vll_rms_LV_BC_pu"].min(),
        #         df["PCU2_Vll_rms_LV_CA_pu"].min()  # Assuming you meant phase C, not repeating phase B
        #     )
        #     max_val = max(
        #         df["VMEAS_LCL"].max(),
        #         df["PCU2_Vll_rms_LV_AB_pu"].max(),
        #         df["PCU2_Vll_rms_LV_BC_pu"].max(),
        #         df["PCU2_Vll_rms_LV_CA_pu"].max()
        #     )
        # else:
        min_val = df["V_inv_meter_rms"].min()
        max_val = df["V_inv_meter_rms"].max()
        ax.set_ylim(min_val-0.05,max_val+0.05)

        # [3,1] Pos Sequence Active Current (pu)
        
        # axes[3,1].plot(df.index, df["PCU1_Idref_inv_pos_pu"], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[3,1].plot(df.index, df["Id_pos_LV"], color=col_pallet[1], label='MEAS',linewidth=1.5)
        axes[3,1].xaxis.set_label_position('top')
        axes[3,1].set_xlabel("Pos Sequence Active Current (pu)",loc="left")
        axes[3,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[3,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[3,1]
        ax.set_facecolor((0.98,0.98,0.98))
        # min_val = min(df["PCU1_Idref_inv_pos_pu"].min(),df["Id_pos_LV"].min())
        # max_val = max(df["PCU1_Idref_inv_pos_pu"].max(),df["Id_pos_LV"].max())
        min_val = df["Id_pos_LV"].min()
        max_val = df["Id_pos_LV"].max()
        ax.set_ylim(min_val-0.1,max_val+0.1)

        #[4,1] Pos Sequence Reactive Current (pu)
        scaling = [1.0,1.0]

        # axes[4,1].plot(df.index, df["PCU2_Iqref_inv_pos_pu"]*scaling[0], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[4,1].plot(df.index, df["Iq_pos_LV"]*scaling[1], color=col_pallet[1], label='MEAS',linewidth=1.5)

        axes[4,1].xaxis.set_label_position('top')
        axes[4,1].set_xlabel("Pos Sequence Reactive Current (pu)",loc="left")
        axes[4,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[4,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[4,1]
        ax.set_facecolor((0.98,0.98,0.98))
        # min_val = min((df["PCU2_Iqref_inv_pos_pu"] * scaling[0]).min(),(df["Iq_pos_LV"] * scaling[1]).min())
        # max_val = max((df["PCU2_Iqref_inv_pos_pu"] * scaling[0]).max(),(df["Iq_pos_LV"] * scaling[1]).max())
        min_val = (df["Iq_pos_LV"]*scaling[1]).min()
        max_val = (df["Iq_pos_LV"]*scaling[1]).max()
        ax.set_ylim(min_val-0.1,max_val+0.1)

        #------------------------------------------------------ MISC PLOTS ---------------------------------------------------

        # #[1,2] Main TX OLTC 
        # axes[1,2].plot(df.index, df["OLTC_TAP_NUMBER"], color=col_pallet[0], label='TAP NUMBER',linewidth=1.5)
        # axes[1,2].xaxis.set_label_position('top')

        # axes[1,2].set_xlabel("Main TX OLTC",loc="left")
        # axes[1,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        # axes[1,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        # ax = axes[1,2]
        # ax.set_facecolor((0.98,0.98,0.98))
        # min_val = df["OLTC_TAP_NUMBER"].min()
        # max_val = df["OLTC_TAP_NUMBER"].max()
        # ax.set_ylim(min_val-2,max_val+2) 
        #         #Create shading for OLTC Operating
        # lowerbound = 0.0
        # upperbound = 0.0
        # signal = df["OLTC_OPERATING"]
        # ax.fill_between(df.index,
        #                  min_val,
        #                  max_val,
        #                  where=(signal<upperbound if upperbound is not None else True)&(signal>lowerbound if lowerbound is not None else True),
        #                  alpha=0.5,
        #                  label= "OLTC OPERATING"
        #                  )
        # axes[4,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)

        # [2,2] FRT Flags

        axes[2,2].plot(df.index, df["LVRT_flag"], color=col_pallet[0], label='LVRT',linewidth=1.5)
        axes[2,2].xaxis.set_label_position('top')
        axes[2,2].plot(df.index, df["HVRT_flag"], color=col_pallet[1], label='HVRT',linewidth=1.5)

        axes[2,2].set_xlabel("FRT Flags",loc="left")
        axes[2,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=3)
        axes[2,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[2,2] 
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = min(df["LVRT_flag"].min(),df["HVRT_flag"].min())
        max_val = max(df["LVRT_flag"].max(),df["HVRT_flag"].max())
        ax.set_ylim(min_val-2,max_val+2)

        # [3,2] Frequency (Hz)
        
        axes[3,2].plot(df.index,df["Fpoc_Hz"], color=col_pallet[0], label='Freq',linewidth=1.5)
        axes[3,2].xaxis.set_label_position('top')

        axes[3,2].set_xlabel("Frequency (Hz)",loc="left")
        axes[3,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[3,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        min_val = df['Fpoc_Hz'].min()
        max_val = df['Fpoc_Hz'].max()
        axes[3,2].set_ylim(min_val -2 , max_val +2)
        ax = axes[3,2] 
        ax.set_facecolor((0.98,0.98,0.98))

        #[4,2] Main TX OLTC 
        axes[4,2].plot(df.index, df["Irms_HV_pu"], color=col_pallet[0], label='POC',linewidth=1.5)
        axes[4,2].plot(df.index, df["Irms_LV_pu"], color=col_pallet[1], label='INV',linewidth=1.5)
        axes[4,2].xaxis.set_label_position('top')
    
        axes[4,2].set_xlabel("RMS Current (pu)",loc="left")
        axes[4,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        
        ax = axes[4,2] 
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = min(df["Irms_HV_pu"].min(),df["Irms_LV_pu"].min())
        max_val = max(df["Irms_HV_pu"].max(),df["Irms_LV_pu"].max())
        ax.set_ylim(min_val-0.1,max_val+0.1)


#----------------------------------------------- TABLE ------------------------------------------------------------

        data = []
        if "Plotter Table" in scenario_dict:
            if not pd.isna(scenario_dict["Plotter Table"]):
                try:
                    plotter_table = ast.literal_eval(scenario_dict["Plotter Table"])
                except (ValueError,SyntaxError)as e:
                    print(f"Error in literal_eval: {e}")
                    print("File Name:" ,scenario_dict["File_Name"], "Format of Plotter Table entry is incorrect")
                    plt.tight_layout()
                    plt.savefig(png_path, bbox_inches="tight")
                    plt.savefig(pdf_path, bbox_inches="tight")
                    plt.clf()
                    plt.close()
                    return

                if len(plotter_table) >0:
                    for item in range(len(plotter_table)):
                        try:
                            if plotter_table[item][0] == "X2R":
                                new_key = "X/R"
                                data.append({new_key: plotter_table[item][1]})
                            else:
                                data.append({plotter_table[item][0]: plotter_table[item][1]})
                        except Exception as e:
                            data.append({"Status": "Invalid entry"})
                            continue
                    if "software" in scenario_dict:
                        data.append({"Software":scenario_dict["software"]})
                else:
                    data.append({"Status": "No data provided"})
        else:
            data.append({"Status": "Plotter Table column in Excel not provided"})
            
        ax = fig.add_subplot(5,3,3)
        rows = len(data)
        cols= 2

                
        for index,row in enumerate(data):

            ax.text(x=0, y=len(data)-index-1, s=list(row.keys())[0], va='center', ha='left', weight='bold', fontsize=8)
            ax.text(x=1.05, y=len(data)-index-1, s=list(row.values())[0], va='center', ha='left', fontsize=8)

        for row in range(rows):
            #row gridlines
            ax.plot(
                [0, cols + 1],
                [row -.5, row - .5],
                ls=':',
                lw='.5',
                c='grey'
            )
        #column highlight
        rect = patches.Rectangle(
            (0, -.5),  # bottom left starting position (x,y)
            1.0,  # width
            len(data),  
            ec='none',
            fc='grey',
            alpha=.2,
            zorder=-1
        )
        #header divider
        ax.plot([1, 1], [-.5, rows-0.5], lw='.5', c='black') 
        ax.add_patch(rect)
        ax.axis('off')
              

        plt.tight_layout()
        plt.savefig(png_path, bbox_inches="tight")
        plt.savefig(pdf_path, bbox_inches="tight")
        plt.clf()
        plt.close()
                












