# heywoodBessWANOverlayer
"""
This is the class for the new plotter for heywood bess wide area studies

"""



from typing import Optional, Union, Dict, List, Callable
from pathlib import Path
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import pandas as pd
from matplotlib.table import Cell, Table
from pallet.plotting import Plotter
import ast



class heywoodbessWANOverlayer(Plotter): 
    def __init__(
        self, 
        pre_process_fn: Optional[Callable] = None
    ):
        self.pre_process_fn = pre_process_fn


    def plot_from_df_and_dict(
            self,
            df1: pd.DataFrame,
            df2: pd.DataFrame,
            remove_first_seconds: Optional[float] = 0.0,
            scenario_dict: Optional[Dict] = None,
            png_path: Optional[Union[str, Path]] = None,
            pdf_path: Optional[Union[str, Path]] = None,
    ):
        if self.pre_process_fn is not None:
            #print(df.columns)
            df1 = self.pre_process_fn(df1)
            df2 = self.pre_process_fn(df2)
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

        col_pallet1 = ["#ab427a",
                "#96b23d",
                "#6162b9",
                "#ce9336",
                "#bc6abc",
                "#5aa554",
                "#ba4758",
                "#46c19a",
                "#b95436",
                "#9e8f41"]
        
        col_pallet2 = [
                "#5aa554",
                "#ba4758",
                "#46c19a",
                "#b95436",
                "#9e8f41",
                "#ab427a",
                "#96b23d",
                "#6162b9",
                "#ce9336",
                "#bc6abc"]

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
        min_val_list = []
        max_val_list = []

        for df, col_pallet,signal_label in [[df1,col_pallet1,"IS"], [df2,col_pallet2,"OOS"]]: #this for loop goes over all the subplots twice, adding the signals from one dataframe per pass. The scaling is handled after the for loop to ensure both signals are shown on the plots.
            

            # [0, 0] POC active power
            channel_name = "REDCLIFFS_220KV_V_PU"
            axes_v[1,0].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_v[1,0].xaxis.set_label_position('top')
            axes_v[1,0].set_xlabel("Voltage - RedCliff Bus (pu)",loc="left")
            axes_v[1,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_v[1,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax1_0 = axes_v[1,0]
            ax1_0.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())


            channel_name = "WODONGA_330KV_V_PU"
            axes_v[2,0].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_v[2,0].xaxis.set_label_position('top')
            axes_v[2,0].set_xlabel("Voltage - Wodonga Bus (pu)",loc="left")
            axes_v[2,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_v[2,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax2_0 = axes_v[2,0]
            ax2_0.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())


            channel_name = "BALLARAT_220KV_V_PU"
            axes_v[3,0].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_v[3,0].xaxis.set_label_position('top')
            axes_v[3,0].set_xlabel("Voltage - Ballarat Bus (pu)",loc="left")
            axes_v[3,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_v[3,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax3_0 = axes_v[3,0]
            ax3_0.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "HEYWOOD_500KV_V_PU"
            axes_v[1,1].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_v[1,1].xaxis.set_label_position('top')
            axes_v[1,1].set_xlabel("Voltage - Heywood Bus (pu)",loc="left")
            axes_v[1,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_v[1,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax1_1 = axes_v[1,1]
            ax1_1.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "GLENROWAN_220KV_V_PU"
            axes_v[2,1].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_v[2,1].xaxis.set_label_position('top')
            axes_v[2,1].set_xlabel("Voltage - Glenrowan Bus (pu)",loc="left")
            axes_v[2,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_v[2,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax2_1 = axes_v[2,1]
            ax2_1.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "HAZELWOOD_500KV_V_PU"
            axes_v[3,1].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_v[3,1].xaxis.set_label_position('top')       
            axes_v[3,1].set_xlabel("Voltage - Hazelwood Bus (pu)",loc="left")
            axes_v[3,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_v[3,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax3_1 = axes_v[3,1]
            ax3_1.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "SHTS_220KV_V_PU"
            axes_v[1,2].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_v[1,2].xaxis.set_label_position('top')       
            axes_v[1,2].set_xlabel("Voltage - Shepparton Bus (pu)",loc="left")
            axes_v[1,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_v[1,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax1_2 = axes_v[1,2]
            ax1_2.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "DEDERANG_220KV_V_PU"
            axes_v[2,2].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_v[2,2].xaxis.set_label_position('top')       
            axes_v[2,2].set_xlabel("Voltage - Dederang Bus (pu)",loc="left")
            axes_v[2,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_v[2,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax2_2 = axes_v[2,2]
            ax2_2.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "PINE_LODGE_BESS_POC_V_PU"
            axes_v[3,2].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_v[3,2].xaxis.set_label_position('top')       
            axes_v[3,2].set_xlabel("Voltage - Pine Lodge Bus (pu)",loc="left")
            axes_v[3,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_v[3,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax3_2 = axes_v[3,2]
            ax3_2.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())
            
        for index, ax in enumerate([ax1_0,ax2_0,ax3_0,ax1_1,ax2_1,ax3_1,ax1_2,ax2_2,ax3_2]):
            min_val=min(min_val_list[index],min_val_list[index+9])
            max_val=max(max_val_list[index],max_val_list[index+9])
            ax.set_ylim(min_val - abs(min_val-max_val)*0.05,max_val+abs(min_val-max_val)*0.05)
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
        min_val_list = []
        max_val_list = []

        for df, col_pallet,signal_label in [[df1,col_pallet1,"IS"], [df2,col_pallet2,"OOS"]]:
        
            # [0, 0] POC active power
            channel_name = "MURRAY_UTUMUT_P_MW"
            axes_l[1,0].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_l[1,0].xaxis.set_label_position('top')
            axes_l[1,0].set_xlabel("Active Power (MW): Murray - Upper Tumut Line (VIC-NSW)",loc="left")
            axes_l[1,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_l[1,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax1_0 = axes_l[1,0]
            ax1_0.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "MURRAY_LTUMUT_P_MW"
            axes_l[2,0].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_l[2,0].xaxis.set_label_position('top')
            axes_l[2,0].set_xlabel("Active Power (MW): Murray - Lower Tumut Line (VIC-NSW)",loc="left")
            axes_l[2,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_l[2,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax2_0 = axes_l[2,0]
            ax2_0.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "WODONGA_JINDERA_P_MW"
            axes_l[3,0].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_l[3,0].xaxis.set_label_position('top')
            axes_l[3,0].set_xlabel("Active Power (MW): Wodonga - Jindera Line (VIC-NSW)",loc="left")
            axes_l[3,0].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_l[3,0].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax3_0 = axes_l[3,0]
            ax3_0.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "REDCLIFFS_BURONGA1_P_MW"
            axes_l[1,1].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_l[1,1].xaxis.set_label_position('top')
            axes_l[1,1].set_xlabel("Active Power (MW): Redcliffs - Buronga 1 Line (VIC-NSW)",loc="left")
            axes_l[1,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_l[1,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax1_1 = axes_l[1,1]
            ax1_1.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "REDCLIFFS_BURONGA2_P_MW"
            axes_l[2,1].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_l[2,1].xaxis.set_label_position('top')
            axes_l[2,1].set_xlabel("Active Power (MW): Redcliffs - Buronga 2 Line (VIC-NSW)",loc="left")
            axes_l[2,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_l[2,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax2_1 = axes_l[2,1]
            ax2_1.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "HEYWOOD_SOUTHEAST1_P_MW"
            axes_l[3,1].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_l[3,1].xaxis.set_label_position('top')       
            axes_l[3,1].set_xlabel("Active Power (MW): Heywood South East Line 1 (SA-VIC)",loc="left")
            axes_l[3,1].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_l[3,1].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax3_1 = axes_l[3,1]
            ax3_1.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "HEYWOOD_SOUTHEAST2_P_MW"
            axes_l[1,2].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_l[1,2].xaxis.set_label_position('top')       
            axes_l[1,2].set_xlabel("Active Power (MW): Heywood South East Line 2 (SA-VIC)",loc="left")
            axes_l[1,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_l[1,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax1_2 = axes_l[1,2]
            ax1_2.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "REDCLIFFS_MONASH_P_MW"
            axes_l[2,2].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_l[2,2].xaxis.set_label_position('top')       
            axes_l[2,2].set_xlabel("Active Power (MW): Murray-link DC Line (VIC-SA)",loc="left")
            axes_l[2,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_l[2,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax2_2 = axes_l[2,2]
            ax2_2.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

            channel_name = "LOYANG_GTOWN_P_MW"
            axes_l[3,2].plot(df.index, df[channel_name], color=col_pallet[0], label=signal_label,linestyle="solid",linewidth=1.5)
            axes_l[3,2].xaxis.set_label_position('top')       
            axes_l[3,2].set_xlabel("Active Power (MW): Basslink Line (VIC-TAS)",loc="left")
            axes_l[3,2].legend(frameon = False, loc = 'lower right',bbox_to_anchor=(1.0,1.0),fontsize = 7, ncol=2)
            axes_l[3,2].grid(True, linestyle=':', which = 'major', alpha=1,linewidth=1.2)
            ax3_2 = axes_l[3,2]
            ax3_2.set_facecolor((0.98,0.98,0.98,0.8))
            min_val_list.append(df[channel_name].min())
            max_val_list.append(df[channel_name].max())

        for index, ax in enumerate([ax1_0,ax2_0,ax3_0,ax1_1,ax2_1,ax3_1,ax1_2,ax2_2,ax3_2]):
            min_val=min(min_val_list[index],min_val_list[index+9])
            max_val=max(max_val_list[index],max_val_list[index+9])
            ax.set_ylim(min_val - abs(min_val-max_val)*0.05,max_val+abs(min_val-max_val)*0.05)


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