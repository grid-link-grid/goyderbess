"""
This is the class for the new plotter for heywoodbess wide area studies

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



class heywoodbessWANPlotter(Plotter): 
    def __init__(
        self, 
        pre_process_fn: Optional[Callable] = None
    ):
        self.pre_process_fn = pre_process_fn


    def plot_from_df_and_dict(
            self,
            df: pd.DataFrame,
            remove_first_seconds: Optional[float] = 0.0,
            scenario_dict: Optional[Dict] = None,
            png_path: Optional[Union[str, Path]] = None,
            pdf_path: Optional[Union[str, Path]] = None,
    ):
       
        print("===> Entering WAN plotter...")
        if self.pre_process_fn is not None:
            #print(df.columns)
            df = self.pre_process_fn(df)
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

        # df = df[df.index > remove_first_seconds]
        # df.index = df.index - remove_first_seconds


#Define Figure for wide area bus voltages

# -------------------------------------------------SUBPLOT CREATION FOR BUS VOLTS-----------------------------------------
        num_rows = 4
        num_cols = 3
        fig_v, axes_v = plt.subplots(num_rows, num_cols, figsize=(16, 10),gridspec_kw={'height_ratios':[1.25, 3, 3, 3]},squeeze = False, sharex=True)
        cm = 1 / 2.54
        fig_v.set_size_inches(42 * cm, 29.7 * cm)  # A4 Size Landscape.
        fig_v.suptitle(scenario_dict["File_Name"],fontweight = 'bold')
        #gs_v = gridspec.GridSpec(num_rows, num_cols, height_ratios=[0.75, 3, 3, 3])
        fig_v.subplots_adjust(left=0.05, right=0.95, bottom=0.05, top=0.95, wspace=0.2, hspace=0.3)
        fig_v.patch.set_linewidth(12)
        # axes[0,0].set_title("POC",loc = 'center',fontsize=13, pad=12)
        # axes[0,1].set_title("INV", loc = 'center',fontsize=13, pad=12)
        # axes[1,2].set_title('MISC',loc='center',fontsize=13, pad=12)
        # axes[0][2].axis('off')

        # [0, 0] POC active power
        channel_name = "REDCLIFFS_220KV_V_PU"
        axes_v[1,0].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_v[1,0].xaxis.set_label_position('top')
        axes_v[1,0].set_xlabel("Voltage - RedCliff Bus (pu)",loc="left")
        axes_v[1,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_v[1,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_v[1,0]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "SYDHAM_500A_V_PU"
        axes_v[2,0].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_v[2,0].xaxis.set_label_position('top')
        axes_v[2,0].set_xlabel("Voltage - Sydenham Bus (pu)",loc="left")
        axes_v[2,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_v[2,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_v[2,0]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "BALLARAT_220KV_V_PU"
        axes_v[3,0].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_v[3,0].xaxis.set_label_position('top')
        axes_v[3,0].set_xlabel("Voltage - Ballarat Bus (pu)",loc="left")
        axes_v[3,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_v[3,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_v[3,0]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "HEYWOOD_500KV_V_PU"
        axes_v[1,1].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_v[1,1].xaxis.set_label_position('top')
        axes_v[1,1].set_xlabel("Voltage - Heywood Bus (pu)",loc="left")
        axes_v[1,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_v[1,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_v[1,1]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "GLENROWAN_220KV_V_PU"
        axes_v[2,1].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_v[2,1].xaxis.set_label_position('top')
        axes_v[2,1].set_xlabel("Voltage - Glenrowan Bus (pu)",loc="left")
        axes_v[2,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_v[2,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_v[2,1]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "HAZELWOOD_500KV_V_PU"
        axes_v[3,1].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_v[3,1].xaxis.set_label_position('top')       
        axes_v[3,1].set_xlabel("Voltage - Hazelwood Bus (pu)",loc="left")
        axes_v[3,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_v[3,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_v[3,1]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "SHTS_220KV_V_PU"
        axes_v[1,2].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_v[1,2].xaxis.set_label_position('top')       
        axes_v[1,2].set_xlabel("Voltage - Shepparton Bus (pu)",loc="left")
        axes_v[1,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_v[1,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_v[1,2]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "DEDERANG_220KV_V_PU"
        axes_v[2,2].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_v[2,2].xaxis.set_label_position('top')       
        axes_v[2,2].set_xlabel("Voltage - Dederang Bus (pu)",loc="left")
        axes_v[2,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_v[2,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_v[2,2]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "SOUTEA_275C_V_PU"
        axes_v[3,2].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_v[3,2].xaxis.set_label_position('top')       
        axes_v[3,2].set_xlabel("Voltage - South East Bus (pu)",loc="left")
        axes_v[3,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_v[3,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_v[3,2]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

# ###



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
            
        #ax = fig.add_subplot(5,3,3)
        ax_v = fig_v.add_axes([0.02, 0.87, 0.2, 0.1])
        rows = len(data)
        cols= 2

                
        for index,row in enumerate(data):

            ax_v.text(x=0, y=len(data)-index-1, s=list(row.keys())[0], va='center', ha='left', weight='bold', fontsize=8)
            ax_v.text(x=1.05, y=len(data)-index-1, s=list(row.values())[0], va='center', ha='left', fontsize=8)

        for row in range(rows):
            #row gridlines
            ax_v.plot(
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
        ax_v.plot([1, 1], [-.5, rows-0.5], lw='.5', c='black') 
        ax_v.add_patch(rect)
        ax_v.axis('off')
        axes_v[0][2].axis('off') #CS removed the top right corner axis
        axes_v[0][1].axis('off')
        axes_v[0][0].axis('off')
              

        plt.tight_layout()
        plt.savefig(png_path.replace(".png","_bus_voltages.png"), bbox_inches="tight")
        plt.savefig(pdf_path.replace(".pdf","_bus_voltages.pdf"), bbox_inches="tight")
        plt.clf()
        plt.close()
                



#Define Figure for wide area interconnector flows

# --------------------------------------SUBPLOT CREATION FOR INTERCONNECTOR FLOWS-----------------------------------
        num_rows = 4
        num_cols = 3
        fig_l, axes_l = plt.subplots(num_rows, num_cols, figsize=(16, 10),gridspec_kw={'height_ratios':[1.25, 3, 3, 3]},squeeze = False, sharex=True)
        cm = 1 / 2.54
        fig_l.set_size_inches(42 * cm, 29.7 * cm)  # A4 Size Landscape.
        fig_l.suptitle(scenario_dict["File_Name"],fontweight = 'bold')
        #gs_v = gridspec.GridSpec(num_rows, num_cols, height_ratios=[0.75, 3, 3, 3])
        fig_l.subplots_adjust(left=0.05, right=0.95, bottom=0.05, top=0.95, wspace=0.2, hspace=0.3)
        fig_l.patch.set_linewidth(12)
        # axes[0,0].set_title("POC",loc = 'center',fontsize=13, pad=12)
        # axes[0,1].set_title("INV", loc = 'center',fontsize=13, pad=12)
        # axes[1,2].set_title('MISC',loc='center',fontsize=13, pad=12)
        # axes[0][2].axis('off')

        # [0, 0] POC active power
        channel_name = "MURRAY_UTUMUT_P_MW"
        axes_l[1,0].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_l[1,0].xaxis.set_label_position('top')
        axes_l[1,0].set_xlabel("Active Power (MW): Murray - Upper Tumut Line (VIC-NSW)",loc="left")
        axes_l[1,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.03),fontsize = 7, ncol=2)
        axes_l[1,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[1,0]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "MURRAY_LTUMUT_P_MW"
        axes_l[2,0].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_l[2,0].xaxis.set_label_position('top')
        axes_l[2,0].set_xlabel("Active Power (MW): Murray - Lower Tumut Line (VIC-NSW)",loc="left")
        axes_l[2,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.03),fontsize = 7, ncol=2)
        axes_l[2,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[2,0]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "WODONGA_JINDERA_P_MW"
        axes_l[3,0].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_l[3,0].xaxis.set_label_position('top')
        axes_l[3,0].set_xlabel("Active Power (MW): Wodonga - Jindera Line (VIC-NSW)",loc="left")
        axes_l[3,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.03),fontsize = 7, ncol=2)
        axes_l[3,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[3,0]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "REDCLIFFS_BURONGA1_P_MW"
        axes_l[1,1].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_l[1,1].xaxis.set_label_position('top')
        axes_l[1,1].set_xlabel("Active Power (MW): Redcliffs - Buronga 1 Line (VIC-NSW)",loc="left")
        axes_l[1,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.03),fontsize = 7, ncol=2)
        axes_l[1,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[1,1]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "REDCLIFFS_BURONGA2_P_MW"
        axes_l[2,1].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_l[2,1].xaxis.set_label_position('top')
        axes_l[2,1].set_xlabel("Active Power (MW): Redcliffs - Buronga 2 Line (VIC-NSW)",loc="left")
        axes_l[2,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.03),fontsize = 7, ncol=2)
        axes_l[2,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[2,1]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "HEYWOOD_SOUTHEAST1_P_MW"
        axes_l[3,1].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_l[3,1].xaxis.set_label_position('top')       
        axes_l[3,1].set_xlabel("Active Power (MW): Heywood South East Line 1 (SA-VIC)",loc="left")
        axes_l[3,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.03),fontsize = 7, ncol=2)
        axes_l[3,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[3,1]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "HEYWOOD_SOUTHEAST2_P_MW"
        axes_l[1,2].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_l[1,2].xaxis.set_label_position('top')       
        axes_l[1,2].set_xlabel("Active Power (MW): Heywood South East Line 2 (SA-VIC)",loc="left")
        axes_l[1,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.03),fontsize = 7, ncol=2)
        axes_l[1,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[1,2]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "REDCLIFFS_MONASH_P_MW"
        axes_l[2,2].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_l[2,2].xaxis.set_label_position('top')       
        axes_l[2,2].set_xlabel("Active Power (MW): Murray-link DC Line (VIC-SA)",loc="left")
        axes_l[2,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.03),fontsize = 7, ncol=2)
        axes_l[2,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[2,2]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))

        channel_name = "LOYANG_GTOWN_P_MW"
        axes_l[3,2].plot(df.index, df[channel_name], color=col_pallet[0], label=channel_name,linestyle="solid",linewidth=1.5)
        axes_l[3,2].xaxis.set_label_position('top')       
        axes_l[3,2].set_xlabel("Active Power (MW): Basslink Line (VIC-TAS)",loc="left")
        axes_l[3,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.03),fontsize = 7, ncol=2)
        axes_l[3,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[3,2]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = df[channel_name].min()
        max_val = df[channel_name].max()
        ax.set_ylim(min_val-0.05*abs(min_val),max_val+0.05*abs(max_val))



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
            
        #ax = fig.add_subplot(5,3,3)
        ax_l = fig_l.add_axes([0.02, 0.87, 0.2, 0.1])
        rows = len(data)
        cols= 2

                
        for index,row in enumerate(data):

            ax_l.text(x=0, y=len(data)-index-1, s=list(row.keys())[0], va='center', ha='left', weight='bold', fontsize=8)
            ax_l.text(x=1.05, y=len(data)-index-1, s=list(row.values())[0], va='center', ha='left', fontsize=8)

        for row in range(rows):
            #row gridlines
            ax_l.plot(
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
        ax_l.plot([1, 1], [-.5, rows-0.5], lw='.5', c='black') 
        ax_l.add_patch(rect)
        ax_l.axis('off')
        axes_l[0][2].axis('off') #CS removed the top right corner axis
        axes_l[0][1].axis('off')
        axes_l[0][0].axis('off')
              

        plt.tight_layout()
        plt.savefig(png_path.replace(".png","_interconnectors.png"), bbox_inches="tight")
        plt.savefig(pdf_path.replace(".pdf","_interconnectors.pdf"), bbox_inches="tight")
        plt.clf()
        plt.close()





#Define Figure for wide area generator specific figure

# --------------------------------------SUBPLOT CREATION FOR GEN SPECIFIC FIGURE-----------------------------------
        num_rows = 4
        num_cols = 3
        fig_l, axes_l = plt.subplots(num_rows, num_cols, figsize=(16, 10),gridspec_kw={'height_ratios':[1.25, 3, 3, 3]},squeeze = False, sharex=True)
        cm = 1 / 2.54
        fig_l.set_size_inches(42 * cm, 29.7 * cm)  # A4 Size Landscape.
        fig_l.suptitle(scenario_dict["File_Name"],fontweight = 'bold')
        #gs_v = gridspec.GridSpec(num_rows, num_cols, height_ratios=[0.75, 3, 3, 3])
        fig_l.subplots_adjust(left=0.05, right=0.95, bottom=0.05, top=0.95, wspace=0.2, hspace=0.3)
        fig_l.patch.set_linewidth(12)
        # axes[0,0].set_title("POC",loc = 'center',fontsize=13, pad=12)
        # axes[0,1].set_title("INV", loc = 'center',fontsize=13, pad=12)
        # axes[1,2].set_title('MISC',loc='center',fontsize=13, pad=12)
        # axes[0][2].axis('off')

        # POC Voltage (pu)
        # axes_l[1,0].plot(df.index, df["V_POC_REF_PU"], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes_l[1,0].plot(df.index, df["V_POC_PU"], color=col_pallet[1], label='MEAS',linewidth=1.5, alpha =0.85)
        axes_l[1,0].xaxis.set_label_position('top')
     
        axes_l[1,0].set_xlabel("POC Voltage (pu)",loc="left")
        axes_l[1,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_l[1,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[1,0]
        ax.set_facecolor((0.98,0.98,0.98))   
        min_val = min(df["V_POC_REF_PU"].min(),df["V_POC_PU"].min())
        max_val = max(df["V_POC_REF_PU"].max(),df["V_POC_PU"].max())
        ax.set_ylim(min_val-0.012,max_val+0.012)

        # POC Active Power (MW)
        scaling = [1.0,1.0]
        axes_l[2,0].plot(df.index, df["PPOC_REF_MW"] * scaling[0], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes_l[2,0].xaxis.set_label_position('top')
        axes_l[2,0].plot(df.index, df["PPOC_MW"] * scaling[1], color=col_pallet[1], label='MEAS',linewidth=1.5, alpha =0.85)

        axes_l[2,0].set_xlabel("# POC Active Power (MW)",loc="left")
        axes_l[2,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_l[2,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[2,0]
        ax.set_facecolor((0.98,0.98,0.98,0.8))
        min_val = min(
            (df["PPOC_REF_MW"] * scaling[0]).min(),
            (df["PPOC_MW"] * scaling[1]).min()
            )
        max_val = max(
            (df["PPOC_REF_MW"] * scaling[0]).max(),
            (df["PPOC_MW"] * scaling[1]).max()
            )
        ax.set_ylim(min_val-2,max_val+2)

        # POC Reactive Power (MVAr)
        scaling = [1.0,1.0]
        axes_l[3,0].plot(df.index, df["QPOC_REF_MVAR"]*scaling[0], color=col_pallet[0], label='REF',linestyle="dashed",linewidth=1.5)
        axes_l[3,0].xaxis.set_label_position('top')
        axes_l[3,0].plot(df.index, df["QPOC_MVAR"]*scaling[1], color=col_pallet[1], label='MEAS',linewidth=1.5, alpha =0.85)

        axes_l[3,0].set_xlabel("POC Reactive Power (MVAr)",loc="left")
        axes_l[3,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_l[3,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[3,0]
        ax.set_facecolor((0.98,0.98,0.98))   
        min_val = min(
            (df["QPOC_REF_MVAR"] * scaling[1]).min(),
            (df["QPOC_MVAR"] * scaling[0]).min()
        )
        max_val = max(
            (df["QPOC_REF_MVAR"] * scaling[1]).max(),
            (df["QPOC_MVAR"] * scaling[0]).max()
        )
        ax.set_ylim(min_val-2,max_val+2)

        # Voltage (pu)
        axes_l[1,1] .plot(df.index, df["V_HYBESS_PU"], color=col_pallet[0], label='INV1 Vrms',linewidth=1.5)
        axes_l[1,1] .xaxis.set_label_position('top')

        axes_l[1,1] .set_xlabel("Voltage (pu)",loc="left")
        axes_l[1,1] .legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_l[1,1] .grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[1,1] 
        ax.set_facecolor((0.98,0.98,0.98))
        min_val = df["V_HYBESS_PU"].min()
        max_val = df["V_HYBESS_PU"].max()
        ax.set_ylim(min_val-0.05,max_val+0.05)

        # Active Power (MW)
        scaling = [105.8,105.8]
        ###
        axes_l[2,1].plot(df.index, df["PREF_MW"]*scaling[0], color=col_pallet[0], label='REF INV1',linestyle="dashed",linewidth=1.5)
        axes_l[2,1].xaxis.set_label_position('top')
        axes_l[2,1].plot(df.index, df["PMEAS_INV_PU"]*scaling[1], color=col_pallet[2], label='MEAS INV1',linewidth=1.5, alpha =0.85)


        axes_l[2,1].set_xlabel("Active Power (MW)",loc="left")
        axes_l[2,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_l[2,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[2,1]
        ax.set_facecolor((0.98,0.98,0.98))
        # min_val = min((df["INV1_PMEAS_PU"]).min()*scaling[0],df["INV2_PMEAS_PU"].min()*scaling[1],df["PREF_MW_INV"].min()*scaling[2])
        # max_val = max(df["INV1_PMEAS_PU"].max()*scaling[0],df["INV2_PMEAS_PU"].max()*scaling[1],df["PREF_MW_INV"].min()*scaling[2])
        min_val = min(
            (df["PREF_MW"] * scaling[0]).min(),
            (df["PMEAS_INV_PU"] * scaling[1]).min(),
        )
        max_val = max(
            (df["PREF_MW"] * scaling[0]).max(),
            (df["PMEAS_INV_PU"] * scaling[1]).max(),
        )
        ax.set_ylim(min_val-2,max_val+2)

        # Reactive Power (MVAr)
        scaling= [63.48,63.48,105.8]
        ####
        axes_l[3,1].plot(df.index, df["QINV_REF_MVAR"]*scaling[0], color=col_pallet[0], label='REF INV1',linestyle="dashed",linewidth=1.5)
        axes_l[3,1].xaxis.set_label_position('top')
        axes_l[3,1].plot(df.index, df["QMEAS_INV_PU"]*scaling[2], color=col_pallet[2], label='MEAS INV1',linewidth=1.5, alpha =0.85)

        axes_l[3,1].set_xlabel("Reactive Power (MVAr)",loc="left")
        axes_l[3,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_l[3,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[3,1]
        ax.set_facecolor((0.98,0.98,0.98))
        # min_val = min((df["INV1_QMEAS_PU"] * scaling[0]).min(),df(["INV2_QMEAS_PU"] * scaling[1]).min(),df(["QREF_MVAR_INV"] * scaling[2]).min())
        min_val = min(
            (df["QINV_REF_MVAR"] * scaling[0]).min(),
            (df["QMEAS_INV_PU"] * scaling[2]).min(),
        )
        max_val = max(
            (df["QINV_REF_MVAR"] * scaling[0]).max(),
            (df["QMEAS_INV_PU"] * scaling[2]).max(),
        )
        # max_val = max((df["INV1_QMEAS_PU"] * scaling[0]).max(),(df["INV2_QMEAS_PU"] * scaling[1]).max(),(df["QREF_MVAR_INV"] * scaling[2]).min())
        ax.set_ylim(min_val-2,max_val+2)


        # Frequency (Hz)
        axes_l[1,2].plot(df.index,df["Hz_POI"], color=col_pallet[0], label='POC',linewidth=1.5)
        axes_l[1,2].xaxis.set_label_position('top')

        axes_l[1,2].set_xlabel("Frequency (Hz)",loc="left")
        axes_l[1,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_l[1,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        freq = df["Hz_POI"]
        min_val = freq.min()
        max_val = freq.max()
        axes_l[1,2].set_ylim(min_val -2 , max_val +2)
        ax = axes_l[1,2] 
        ax.set_facecolor((0.98,0.98,0.98))

        # Active Current (pu)
        axes_l[2,2].xaxis.set_label_position('top')
        # axes_l[2,2].plot(df.index, df["INV1_ID_COMMAND_PU"], color=col_pallet[0], label="REF INV1", linestyle = 'dashed', linewidth=1.5)
        axes_l[2,2].plot(df.index, df["INV1_Id_PU"], color=col_pallet[1], label='MEAS INV1',linewidth=1.5, alpha =0.85)

        axes_l[2,2].set_xlabel("Active Current (pu)",loc="left")
        axes_l[2,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 6, ncol=2)
        axes_l[2,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[2,2]
        ax.set_facecolor((0.98,0.98,0.98))
        # min_val = min(df["INV1_ID_COMMAND_PU"].min(),df["INV1_Id_PU"].min())
        # max_val = max(df["INV1_ID_COMMAND_PU"].max(),df["INV1_Id_PU"].max())
        min_val = df["INV1_Id_PU"].min()
        max_val = df["INV1_Id_PU"].max()
        ax.set_ylim(min_val-0.1,max_val+0.1)

        # Reactive Current (pu)
        axes_l[3,2].xaxis.set_label_position('top')
        # axes_l[3,2].plot(df.index, df["INV1_VUEL_PU"], color=col_pallet[0], label='REF INV1', linestyle = 'dashed', linewidth=1.5)
        axes_l[3,2].plot(df.index, df["INV1_Iq_PU"], color=col_pallet[1], label='MEAS INV1',linewidth=1.5, alpha =0.85)


        axes_l[3,2].set_xlabel("Reactive Current (pu)",loc="left")
        axes_l[3,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes_l[3,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes_l[3,2]
        ax.set_facecolor((0.98,0.98,0.98))
        # min_val = min(df["INV1_VUEL_PU"].min(),df["INV1_Iq_PU"].min())
        # max_val = max(df["INV1_VUEL_PU"].max(),df["INV1_Iq_PU"].max())
        min_val = df["INV1_Iq_PU"].min()
        max_val = df["INV1_Iq_PU"].max()
        # min_val = df["IQ_MEAS"].min()
        # max_val = df["IQ_MEAS"].max()
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
            
        #ax = fig.add_subplot(5,3,3)
        ax_l = fig_l.add_axes([0.02, 0.87, 0.2, 0.1])
        rows = len(data)
        cols= 2

                
        for index,row in enumerate(data):

            ax_l.text(x=0, y=len(data)-index-1, s=list(row.keys())[0], va='center', ha='left', weight='bold', fontsize=8)
            ax_l.text(x=1.05, y=len(data)-index-1, s=list(row.values())[0], va='center', ha='left', fontsize=8)

        for row in range(rows):
            #row gridlines
            ax_l.plot(
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
        ax_l.plot([1, 1], [-.5, rows-0.5], lw='.5', c='black') 
        ax_l.add_patch(rect)
        ax_l.axis('off')
        axes_l[0][2].axis('off') #CS removed the top right corner axis
        axes_l[0][1].axis('off')
        axes_l[0][0].axis('off')
              
        print("===> Completed plotting, saving figures...")
        plt.tight_layout()
        plt.savefig(png_path.replace(".png","_genSpecific.png"), bbox_inches="tight")
        plt.savefig(pdf_path.replace(".pdf","_genSpecific.pdf"), bbox_inches="tight")
        plt.clf()
        plt.close()








