"""
Description of code: This code creates the class LmthPlotterPsse, which is the new plotter for the learmonth BESS simulation outputs using PSSE
This class is specific for the learmonth BESS psse outputs required for DMAT. 
This code is to be used with 1_run_psse_studies.py or folder_iteration_psse.py

Note:
The plotter requires that there is column in the excel spec sheet named "Plotter Table" in order to populate the table. Otherwise, there will be a table with no relevant information provided


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



class goyderbessBenchmarkPlotter(): 
   
    def plot_benchmarking(
            self,
            study_name: str,
            error_studies: str,
            df_psse: pd.DataFrame,
            df_pscad: pd.DataFrame,
            scenario_dict: Optional[Dict] = None,
            png_path: Optional[Union[str, Path]] = None,
            pdf_path: Optional[Union[str, Path]] = None,
    ):
        # if self.pre_process_fn is not None:
        #     print(df.columns)
        #     subdir = os.path.dirname(png_path)
        #     df, scenario_dict = self.pre_process_fn(df,scenario_dict,subdir)



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
        
        #Col Pallet with less green and red - For luke - review
        # col_pallet = [
        #     "#00429d",  # Deep Blue (kept, good contrast)
        #     "#ffb703",  # Bright Amber (contrast with Deep Blue)
        #     "#118ab2",  # Cyan (kept, suitable for accessibility)
        #     "#264653",  # Dark Blue-Green (kept, works well for contrast)
        #     "#6a4c93",  # Muted Purple (accessible and distinct)
        #     "#ff6700",  # Vivid Orange (replaces mustard yellow for contrast)
        #     "#2a9d8f",  # Teal (good for accessibility)
        #     "#e76f51",  # Coral Red (replaces bright yellow, less ambiguous)
        #     "#6c757d",  # Gray-Blue (kept, neutral tone)
        #     "#dda15e"   # Warm Gold (replaces yellow gold, accessible)
        # ]


# -------------------------------------------------SUBPLOT CREATION ------------------------------------------------------
        fig, axes = plt.subplots(2,2, figsize=(10, 12),squeeze = False, sharex=True)
        cm = 1 / 2.54
        fig.set_size_inches(42 * cm, 29.7 * cm)  # A4 Size Landscape.
        #fig.suptitle(scenario_dict["File_Name"],fontweight = 'bold')
        fig.subplots_adjust(left=0.05, right=0.95, bottom=0.05, top=0.9, wspace=0.2, hspace=0.3)
  

#------------------- ---------------------------------BENCHMARK PLOTS -----------------------------------------------------------
        # [0, 0] POC active power
        axes[0,0].plot(df_pscad.index, df_pscad["PLANT_P_HV"], color=col_pallet[0], label='POC PSCAD',linewidth=1.5)
        axes[0,0].xaxis.set_label_position('top')
        axes[0,0].plot(df_psse.index, df_psse["PSSE_PPOC_MW"], color=col_pallet[1], label='POC PSSE',linewidth=1.5)
        axes[0,0].set_xlabel("Active Power (MW)",loc="left")
        axes[0,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[0,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[0,0]
        ax.set_facecolor((0.98,0.98,0.98,0.8))

        if study_name in error_studies:
            axes[0,0].plot(df_pscad.index, df_pscad["MW_UPPER"], color='red',linewidth=1.5,linestyle='dashed')
            axes[0,0].plot(df_pscad.index, df_pscad["MW_LOWER"], color='red',linewidth=1.5,linestyle='dashed')
            min_val = min(df_pscad["PLANT_P_HV"].min(),df_pscad["MW_LOWER"].min(),df_psse["PSSE_PPOC_MW"].min())
            max_val = max(df_pscad["PLANT_P_HV"].max(),df_pscad["MW_UPPER"].max(),df_psse["PSSE_PPOC_MW"].max())
            ax.set_ylim(min_val-125,max_val+125)

        else:
            min_val = min(df_pscad["PLANT_P_HV"].min(),df_psse["PSSE_PPOC_MW"].min())
            max_val = max(df_pscad["PLANT_P_HV"].max(),df_psse["PSSE_PPOC_MW"].max())
            ax.set_ylim(min_val-125,max_val+125)


        # [0,1] POC Reactive power
        axes[0,1].plot(df_pscad.index, df_pscad["PLANT_Q_HV"], color=col_pallet[0], label='POC PSCAD',linewidth=1.5)
        axes[0,1].xaxis.set_label_position('top')
        axes[0,1].plot(df_psse.index, df_psse["PSSE_QPOC_MVAR"], color=col_pallet[1], label='POC PSSE',linewidth=1.5)
        axes[0,1].set_xlabel("Reactive Power (MVAr)",loc="left")
        axes[0,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[0,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[0,1]
        ax.set_facecolor((0.98,0.98,0.98))   

        if study_name in error_studies:
            axes[0,1].plot(df_pscad.index, df_pscad["MVAr_UPPER"], color='red',linewidth=1.5,linestyle='dashed')
            axes[0,1].plot(df_pscad.index, df_pscad["MVAr_LOWER"], color='red',linewidth=1.5,linestyle='dashed')
            min_val = min(df_pscad["PLANT_Q_HV"].min(),df_pscad["MVAr_LOWER"].min()) 
            max_val = max(df_pscad["PLANT_Q_HV"].max(),df_pscad["MVAr_UPPER"].max()) 
            ax.set_ylim(min_val-49.375,max_val+49.375)

        else:
            min_val = df_pscad["PLANT_Q_HV"].min()
            max_val = df_pscad["PLANT_Q_HV"].max() 
            ax.set_ylim(min_val-49.375,max_val+49.375)


        # [1,0] POC Voltage  
        axes[1,0].plot(df_pscad.index, df_pscad["PLANT_V_HV"], color=col_pallet[0], label='POC PSCAD',linewidth=1.5)
        axes[1,0].xaxis.set_label_position('top')
        axes[1,0].plot(df_psse.index, df_psse["PSSE_V_POC_PU"], color=col_pallet[1], label='POC PSSE',linewidth=1.5)
        axes[1,0].set_xlabel("Voltage (pu)",loc="left")
        axes[1,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
        axes[1,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
        ax = axes[1,0]
        ax.set_facecolor((0.98,0.98,0.98))
        
        if study_name in error_studies:
            axes[1,0].plot(df_pscad.index, df_pscad["kV_UPPER"], color='red',linewidth=1.5,linestyle='dashed')
            axes[1,0].plot(df_pscad.index, df_pscad["kV_LOWER"], color='red',linewidth=1.5,linestyle='dashed')
            min_val = min(df_pscad["PLANT_V_HV"].min(),df_pscad["kV_LOWER"].min(),df_psse["PSSE_V_POC_PU"].min())
            max_val = max(df_pscad["PLANT_V_HV"].max(),df_pscad["kV_UPPER"].max(),df_psse["PSSE_V_POC_PU"].max())
            ax.set_ylim(min_val-0.05,max_val+0.05)

        else:  
            min_val = min(df_pscad["PLANT_V_HV"].min(),df_psse["PSSE_V_POC_PU"].min())
            max_val = max(df_pscad["PLANT_V_HV"].max(),df_psse["PSSE_V_POC_PU"].max())
            ax.set_ylim(min_val-0.05,max_val+0.05)


        # [1,1] FRT Flags 
        if study_name != 'Fgrid Steps':
            axes[1,1].plot(df_pscad.index, df_pscad["LVRT_FLAG"], color=col_pallet[0], label='PPC PSCAD',linewidth=1.5)
            axes[1,1].plot(df_pscad.index, df_pscad["HVRT_FLAG"], color=col_pallet[1], label='PPC PSCAD',linewidth=1.5)
            axes[1,1].xaxis.set_label_position('top')
            axes[1,1].plot(df_psse.index, df_psse["PSSE_LVRT_FLAG"], color=col_pallet[3], label='PPC PSSE',linewidth=1.5)
            axes[1,1].plot(df_psse.index, df_psse["PSSE_HVRT_FLAG"], color=col_pallet[4], label='PPC PSSE',linewidth=1.5)

            axes[1,1].set_xlabel("FRT Flags",loc="left")
            axes[1,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes[1,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax = axes[1,1]
            ax.set_facecolor((0.98,0.98,0.98))
            # min_val = min(df_pscad["PPC_FRT_FLAG"].min(),df_pscad["INV_FRT_FLAG"].min(),df_psse["PSSE_PPC_FRT_ACTIVE"].min(),
            #               df_psse["PSSE_INV1_FRT_STATE"].min())
            # max_val = max(df_pscad["PPC_FRT_FLAG"].max(),df_pscad["INV_FRT_FLAG"].max(),df_psse["PSSE_PPC_FRT_ACTIVE"].max(),
            #               df_psse["PSSE_INV1_FRT_STATE"].max())
            # ax.set_ylim(min_val-2,max_val+2)

        else:
            scaling = 50 
            offset = 50
            axes[1,1].plot(df_pscad.index,df_pscad["Fpoc_Hz"], color=col_pallet[0], label='PSCAD Freq',linewidth=1.5)
            axes[1,1].xaxis.set_label_position('top')
            axes[1,1].plot(df_psse.index,df_psse["PSSE_POC_F_PU"]*scaling+offset, color=col_pallet[1], label='PSSE Freq',linewidth=1.5)
            
            axes[1,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes[1,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            freq = df_psse["PSSE_POC_F_PU"]*scaling + offset
            min_val = min(freq.min(),df_pscad["Fpoc_Hz"].min())
            max_val = max(freq.max(),df_pscad["Fpoc_Hz"].max())
            ax = axes[1,1] 
            ax.set_facecolor((0.98,0.98,0.98))
            ax.set_ylim(min_val -2 , max_val +2)
        


#----------------------------------------------- TABLE ------------------------------------------------------------

        """
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
                
        """
        plt.tight_layout()
        plt.savefig(png_path, bbox_inches="tight")
        plt.savefig(pdf_path, bbox_inches="tight")
        plt.clf()
        plt.close()











