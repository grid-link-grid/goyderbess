

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



class heywoodbessPscadInitialisation(Plotter): 
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
        axes[0,0].plot(df.index, df["Pref_MW"], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[0,0].xaxis.set_label_position('top')
        axes[0,0].plot(df.index, df["Ppoc_MW"], color=col_pallet[1], label='MEAS',linewidth=1.5)
        ppoc_mw_final = df["Ppoc_MW"].iloc[-1]
        
        axes[0,0].set_xlabel("Active Power (MW)",loc="left")
        axes[0,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[0,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[0,0]
        ax.set_facecolor((0.98,0.98,0.98,0.8))

        min_val = ppoc_mw_final - 5
        max_val = ppoc_mw_final + 5
        ax.set_ylim(min_val,max_val)


        # [1,0] POC Reactive power
        axes[1,0].plot(df.index, df["PwrRtSpntTot_MVAr"], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[1,0].xaxis.set_label_position('top')
        axes[1,0].plot(df.index, df["Qpoc_MVAr"], color=col_pallet[1], label='MEAS',linewidth=1.5)
        qpoc_mvar_final = df["Qpoc_MVAr"].iloc[-1]

        axes[1,0].set_xlabel("Reactive Power (MVAr)",loc="left")
        axes[1,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[1,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[1,0]
        ax.set_facecolor((0.98,0.98,0.98))   
        min_val = qpoc_mvar_final - 5
        max_val = qpoc_mvar_final + 5
        ax.set_ylim(min_val,max_val)

        # [2,0] POC Voltage  
        axes[2,0].plot(df.index, df["Vref_pu"], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[2,0].plot(df.index, df["Vrms_poc_pu"], color=col_pallet[1], label='MEAS',linewidth=1.5)

        vpoc_pu_final = df["Vrms_poc_pu"].iloc[-1]

        axes[2,0].xaxis.set_label_position('top')
        axes[2,0].axhline(y=float(scenario_dict["substitutions"]["NOTE_PPC_HVRT_pu"]), color='black', label='FREEZE THRESHOLD',linestyle="dashed",linewidth=1.5)
        axes[2,0].axhline(y=float(scenario_dict["substitutions"]["NOTE_PPC_LVRT_pu"]), color='black',linestyle="dashdot",linewidth=1.5)

        axes[2,0].set_xlabel("Voltage (pu)",loc="left")
        axes[2,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[2,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[2,0]
        ax.set_facecolor((0.98,0.98,0.98))   

        min_val = vpoc_pu_final - 0.01
        max_val = vpoc_pu_final + 0.01

        ax.set_ylim(min_val,max_val)


        # [3,0] POC Pos Sequence Active Current (pu) 
        axes[3,0].plot(df.index, df["Id_pos_pu"], color=col_pallet[0], label='Id',linewidth=1.5)
        axes[3,0].xaxis.set_label_position('top')

        axes[3,0].set_xlabel("Pos Sequence Active Current (pu)",loc="left")
        axes[3,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[3,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[3,0]
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = df["Id_pos_pu"].min()
        max_val = df["Id_pos_pu"].max()
        ax.set_ylim(min_val-0.1,max_val+0.1)



        # [4,0] POC Pos Sequence Active Current (pu) 
        axes[4,0].plot(df.index, df["Iq_pos_pu"], color=col_pallet[0], label='Iq',linewidth=1.5)
        axes[4,0].xaxis.set_label_position('top')

        axes[4,0].set_xlabel("Pos Sequence Reactive Current (pu)",loc="left")
        axes[4,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[4,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[4,0]
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = df["Iq_pos_pu"].min()
        max_val = df["Iq_pos_pu"].max()
        ax.set_ylim(min_val-0.1,max_val+0.1)

#-----------------------------------------INV PLOTS----------------------------------------------------------------
        # [0,1] INV active Power 
        scaling = [1.0,25.0]
        axes[0,1].plot(df.index, df["MWSpt_SCS"]*scaling[0], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[0,1].xaxis.set_label_position('top')
        axes[0,1].plot(df.index, df["PCU1_P_LV"]*scaling[1], color=col_pallet[1], label='MEAS',linewidth=1.5)


        axes[0,1].set_xlabel("Active Power (MW)",loc="left")
        axes[0,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[0,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[0,1]
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = min(
            (df["MWSpt_SCS"] * scaling[0]).min(),
            (df["PCU1_P_LV"] * scaling[1]).min()
        )
        max_val = max(
            (df["MWSpt_SCS"] * scaling[0]).max(),
            (df["PCU1_P_LV"] * scaling[1]).max()
        )
        ax.set_ylim(min_val-20,max_val+20)



        # [1,1] INV reactive Power 
        scaling = [1.0, 1.0]
        axes[1,1].plot(df.index, df["MVArSpt_SCS"]*scaling[0], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[1,1].xaxis.set_label_position('top')
        axes[1,1].plot(df.index, df["PCU1_Q_LV"]*scaling[1], color=col_pallet[1], label='MEAS',linewidth=1.5)


        axes[1,1].set_xlabel("Reactive Power (MVAr)",loc="left")
        axes[1,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[1,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[1,1]
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = min(
            (df["MVArSpt_SCS"] * scaling[0]).min(),
            (df["PCU1_Q_LV"] * scaling[1]).min()
        )
        max_val = max(
            (df["MVArSpt_SCS"] * scaling[0]).max(),
            (df["PCU1_Q_LV"] * scaling[1]).max()
        )
        ax.set_ylim(min_val-20,max_val+20)

        # [2,1] INV voltage
        axes[2,1] .plot(df.index, df["PCU1_Vrms_LV"], color=col_pallet[0], label='Vrms',linewidth=1.5)
        axes[2,1] .xaxis.set_label_position('top')

        axes[2,1].axhline(y=float(scenario_dict["substitutions"]["NOTE_INV_LVRT_Out_pu"]), color='black', label='FRT THRESHOLD',linestyle="dashed",linewidth=1.5)
        axes[2,1].axhline(y=float(scenario_dict["substitutions"]["NOTE_INV_LVRT_In_pu"]), color='black',linestyle="dashdot",linewidth=1.5)
        axes[2,1].axhline(y=float(scenario_dict["substitutions"]["NOTE_INV_HVRT_Out_pu"]), color='blue',linestyle="dashed",linewidth=1.5)
        axes[2,1].axhline(y=float(scenario_dict["substitutions"]["NOTE_INV_HVRT_In_pu"]), color='blue',linestyle="dashdot",linewidth=1.5)

        axes[2,1] .set_xlabel("Voltage (pu)",loc="left")
        axes[2,1] .legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[2,1] .grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[2,1] 
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = df["PCU1_Vrms_LV"].min()
        max_val = df["PCU1_Vrms_LV"].max()
        ax.set_ylim(min_val-0.05,max_val+0.05)

        #[3,1] Pos Sequence Active Current (pu)
        
        axes[3,1].plot(df.index, df["PCU1_Idref_inv_pos_pu"], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[3,1].plot(df.index, df["PCU1_Id_inv_pos_pu"], color=col_pallet[1], label='MEAS',linewidth=1.5)
        axes[3,1].xaxis.set_label_position('top')
        axes[3,1].set_xlabel("Pos Sequence Active Current (pu)",loc="left")
        axes[3,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[3,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[3,1]
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = min(df["PCU1_Idref_inv_pos_pu"].min(),df["PCU1_Id_inv_pos_pu"].min())
        max_val = max(df["PCU1_Idref_inv_pos_pu"].max(),df["PCU1_Id_inv_pos_pu"].max())
        ax.set_ylim(min_val-0.1,max_val+0.1)

        #[4,1] Pos Sequence Reactive Current (pu)
        scaling = [1.0,1.0]

        axes[4,1].plot(df.index, df["PCU1_Iqref_inv_pos_pu"]*scaling[0], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes[4,1].plot(df.index, df["PCU1_Iq_inv_pos_pu"]*scaling[1], color=col_pallet[1], label='MEAS',linewidth=1.5)

        axes[4,1].xaxis.set_label_position('top')
        axes[4,1].set_xlabel("Pos Sequence Reactive Current (pu)",loc="left")
        axes[4,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[4,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[4,1]
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = min(
                (df["PCU1_Iqref_inv_pos_pu"] * scaling[0]).min(),
                (df["PCU1_Iq_inv_pos_pu"] * scaling[1]).min()
        )
        max_val = max(
                (df["PCU1_Iqref_inv_pos_pu"] * scaling[0]).max(),
                (df["PCU1_Iq_inv_pos_pu"] * scaling[1]).max()
        )
        ax.set_ylim(min_val-0.1,max_val+0.1)

        #------------------------------------------------------ MISC PLOTS ---------------------------------------------------

        #[1,2] Main TX OLTC 
        axes[1,2].plot(df.index, df["OLTC_TAP_NUMBER"], color=col_pallet[0], label='TAP NUMBER',linewidth=1.5)
        axes[1,2].xaxis.set_label_position('top')

        axes[1,2].set_xlabel("Main TX OLTC",loc="left")
        axes[1,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[1,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[1,2]
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = df["OLTC_TAP_NUMBER"].min()
        max_val = df["OLTC_TAP_NUMBER"].max()
        ax.set_ylim(min_val-2,max_val+2) 
                #Create shading for OLTC Operating
        lowerbound = 0.0
        upperbound = 0.0
        signal = df["OLTC_OPERATING"]
        ax.fill_between(df.index,
                         min_val,
                         max_val,
                         where=(signal<upperbound if upperbound is not None else True)&(signal>lowerbound if lowerbound is not None else True),
                         alpha=0.5,
                         label= "OLTC OPERATING"
                         )
        axes[4,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)

        # [2,2] FRT Flags

        axes[2,2].plot(df.index, df["PPC_FRT_FLAG"], color=col_pallet[0], label='PPC',linewidth=1.5)
        axes[2,2].xaxis.set_label_position('top')
        axes[2,2].plot(df.index, df["INV_FRT_FLAG"], color=col_pallet[1], label='INV',linewidth=1.5)

        axes[2,2].set_xlabel("FRT Flags",loc="left")
        axes[2,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=3)
        axes[2,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[2,2] 
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = min(df["PPC_FRT_FLAG"].min(),df["INV_FRT_FLAG"].min())
        max_val = max(df["PPC_FRT_FLAG"].max(),df["INV_FRT_FLAG"].max())
        ax.set_ylim(min_val-2,max_val+2)

        #[3,2] Frequency (Hz)
        axes[3,2].plot(df.index,df["Fslack_meas_Hz"], color=col_pallet[0], label='Freq',linewidth=1.5)
        axes[3,2].xaxis.set_label_position('top')

        axes[3,2].set_xlabel("Frequency (Hz)",loc="left")
        axes[3,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[3,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        min_val = df['Fslack_meas_Hz'].min()
        max_val = df['Fslack_meas_Hz'].max()
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
                            data.append({plotter_table[item][0]: plotter_table[item][1]})
                        except Exception as e:
                            data.append({"Status": "Invalid entry"})
                            continue
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
                












