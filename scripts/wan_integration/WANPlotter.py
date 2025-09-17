from typing import Optional, Union, Dict, List
from pathlib import Path
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import pandas as pd
import numpy as np
from matplotlib.table import Cell, Table
import textwrap as twp
from dataclasses import dataclass
from pallet.plotting import Plotter
import warnings
import math



@dataclass
class title:
    fontsize : float
    loc: str
    
@dataclass
class legend:
    loc : str
    bbox_to_anchor : tuple
    frameon : bool
    propsize: float

@dataclass
class table:
    shading_colour: str
    gridline_colour: str
    header_fontsize: float
    data_fontsize: float

@dataclass
class pen:
    linestyle : str
    linewidth : float
    colour_pallet : list

@dataclass
class axis:
    pens: pen
    title : title
    legend: legend

@dataclass
class figure_style:
    axes : axis
    tables : table

def color_wheel(index, total_items):
    # Use a color wheel approach based on hue
    hue = index / total_items  # Calculate the hue based on position
    color = plt.cm.hsv(hue)  # Use the HSV colormap to generate a color
    return color

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

default_style = figure_style(
    axes = axis(
            pens=pen(
                linestyle='solid',
                linewidth=2,
                colour_pallet = col_pallet
                ),
            title=title(
                fontsize=10,
                loc='left',
                ),
            legend=legend(
                loc="lower right",
                bbox_to_anchor=(1.0, 1.0),
                frameon=False,
                propsize=8,
                ),
            ),
    tables = table(
            shading_colour='grey',
            gridline_colour='grey',
            header_fontsize=9,
            data_fontsize=8,
            ),

    )


warnings.filterwarnings( "ignore", module = "matplotlib\..*" )

EVENT_TABLE_ON = False
FIGURE_BORDER_COLOUR = "#1A5363"
# ["red","blue","green","orange","purple","pink","cyan","black","brown","grey","olive"]

class WANPlotter(Plotter):
    def __init__(
            self,
            plot_definitions: List[Dict],
            remove_first_seconds: float = 0.0,
            style_class: Optional[figure_style] = default_style
            ):
        self.plot_definitions = plot_definitions
        self.remove_first_seconds = remove_first_seconds
        self.style_class = style_class


    def plot_from_df_and_dict(
            self,
            df: pd.DataFrame,
            png_path: Optional[Union[str, Path]] = None,
            pdf_path: Optional[Union[str, Path]] = None,
    ):
        



        plt.clf()
        plt.close()
        matplotlib.use('Agg')

        df_copy = df.copy()


        # Remove the initialisation period
        df_copy = df_copy[df_copy.index > self.remove_first_seconds]
        df_copy.index = df_copy.index - self.remove_first_seconds

        # Check if the specified columns exist in the dataframe
        all_columns = []
        for plot_definition in self.plot_definitions:
            for column in plot_definition["columns"]:
                all_columns.append(column["channel"])


        non_existing_columns = set(all_columns) - set(df_copy.columns)
        if non_existing_columns:
            raise ValueError(f"Columns not found in the DataFrame: {non_existing_columns}")

        # Calculate the number of rows and columns for subplots
        #num_plots = len(self.plot_definitions)
        
        num_rows = 3 
        num_cols = math.ceil(len(self.plot_definitions)/num_rows)
    

        # Create subplots
        fig, axes = plt.subplots(num_rows, num_cols, figsize=(16, 10), squeeze=False,sharex=True,)
        # plt.rcParams['figure.constrained_layout.use'] = True
        # plt.rcParams['figure.autolayout'] = False

        cm = 1 / 2.54
        fig.set_size_inches(42 * cm, 29.7 * cm)  # A4 Size Landscape.
        plt.subplots_adjust(left=0.2, right=0.8, bottom=0.1, top=0.9, wspace=0.1, hspace=0.1)
        plt.ion()
        fig.canvas.draw()
        fig.patch.set_linewidth(12)
        fig.patch.set_edgecolor(FIGURE_BORDER_COLOUR)


        # if "should_trip" in scenario_dict and scenario_dict["should_trip"]:
        #     fig.suptitle(t='SHOULD TRIP',x=0.33,y=0.93, ha='right',va='center',fontsize=10,fontweight='bold', color='red')

        # ax = axes[0,0]
        # ax.set_title('MISC',loc='center',fontsize=13,pad = 22)

        # ax = axes[0,1]
        # ax.set_title('POC',loc='center',fontsize=13,pad = 10)

        # ax = axes[0,2]
        # ax.set_title('INV',loc='center',fontsize=13,pad = 10)


        AXIS_COLOUR = (0, 0, 0)
        SECONDARY_COLOUR = (0.98, 0.98, 0.98)

        col=0
        row=0

        # add plot
        for i, plot_definition in enumerate(self.plot_definitions):

            ax=axes[row][col]
            #add trace
            for col_idx, column in enumerate(plot_definition["columns"]):
                col_name = column["channel"]
                if "alias" in column:
                    alias = column["alias"]
                else:
                    alias = column["channel"]
                scaled_col_name = f"{col_name}_scaled"
                df_copy[scaled_col_name] = df_copy[col_name] * plot_definition["scaling"][col_idx]
                ax.plot(df_copy.index, df_copy[scaled_col_name], color = color_wheel(col_idx,len(plot_definition["columns"])), label=alias, linewidth=self.style_class.axes.pens.linewidth,linestyle=self.style_class.axes.pens.linestyle)
                #for the very first signal the max only needs to consider that signal
                if col_idx == 0:
                    y_max = max(df_copy[scaled_col_name])
                    y_min = min(df_copy[scaled_col_name])

                #for all signals following the first
                else:
                    y_max = max(y_max,max(df_copy[scaled_col_name]))
                    y_min = min(y_min,min(df_copy[scaled_col_name]))

                #account for shading requirememts
                if "shade_vertical" in plot_definition.keys():

                    if plot_definition["shade_vertical"]["enable"]:

                        lowerbound = plot_definition["shade_vertical"]["lowerbound"]
                        upperbound = plot_definition["shade_vertical"]["upperbound"]
                        signal = df_copy[plot_definition["shade_vertical"]["signal"]]

                        ax.fill_between(
                            df_copy.index,
                            min(df_copy[col_name]), 
                            max(df_copy[col_name]),
                            where=(signal<upperbound if upperbound is not None else True)&(signal>lowerbound if lowerbound is not None else True),
                            alpha=0.5,
                            label=plot_definition["shade_vertical"]["alias"])

            row = row+1
            if row >= num_rows:
                col = col+1
                if col ==0:
                    row = 1
                else:
                    row =0


            ax.ticklabel_format(useOffset=False)
            # add h-lines
            # if "h-lines" in plot_definition.keys():
            #     for hline in plot_definition["h-lines"]:
            #         yval = float(scenario_dict["substitutions"][hline["gs"]])
            #         label_alias = None
            #         if "legend" in hline:
            #             if hline["legend"]:
            #                 if "alias" in hline:
            #                     label_alias = hline["alias"]
            #                 else:
            #                     label_alias = hline["gs"]  
            #         ax.axhline(y=yval, color='black', linestyle='-.', lw=0.6, label = label_alias)

            #set y-axis-limits to this by default
            y_mid = (y_max + y_min)/2.0
            y_axis_max = y_max+0.05*(y_max-y_mid)
            y_axis_min = y_min-0.05*(y_max-y_mid)
            ax.set_ylim([y_axis_min,y_axis_max])

            #adjust axis limits
            if "y_axis_min_range" in plot_definition:
                y_range = plot_definition["y_axis_min_range"]
                
                y_axis_max = max(y_max+0.05*(y_max-y_mid),y_mid + 0.5*y_range)
                y_axis_min = min(y_min-0.05*(y_max-y_mid),y_mid - 0.5*y_range)
                ax.set_ylim([y_axis_min,y_axis_max])

            if "y_axis_max" in plot_definition:
                y_axis_max = plot_definition["y_axis_max"]
                ax.set_ylim([ax.get_ylim()[0],y_axis_max])
            if "y_axis_min" in plot_definition:
                y_axis_min = plot_definition["y_axis_min"]
                ax.set_ylim([y_axis_min,ax.get_ylim()[1]])

            ax.xaxis.set_label_position('top')
            ax.set_xlabel(plot_definition["title"],loc=self.style_class.axes.title.loc,fontsize=self.style_class.axes.title.fontsize)
            num_legend_cols = math.ceil(len(plot_definition["columns"])/3)

            if num_legend_cols <4:
                ax.legend(frameon = self.style_class.axes.legend.frameon,loc = self.style_class.axes.legend.loc, bbox_to_anchor=self.style_class.axes.legend.bbox_to_anchor, ncol=num_legend_cols,prop={'size': self.style_class.axes.legend.propsize})
            ax.grid(True)
            
                

        # if EVENT_TABLE_ON:

        #     #Define Scenario Table
        #     table1_data  = [
        #         ["Filename", scenario_dict["File_Name"]],
        #         ["Test No.",scenario_dict["Test No"]],
        #         ["Fault Level [MVA]", scenario_dict["Grid_FL_MVA_sig"]],
        #         ["X2R", scenario_dict["Grid_X2R_sig"]],
        #         ["SCR", scenario_dict["Grid_SCR"]],
        #     ]
        #     #Get number of rows in table prior to events being added.
        #     init_data_rows = len(table1_data)
        #     #Only display events if they exist in scenario file

        #     if "Event_Data" in scenario_dict:
        #         events = scenario_dict["Event_Data"]
        #         num_events = len(events)
                
        #         #Remove empty strings
        #         events = list(filter(None,events))
        #         events_char_limit = []
        #         ROW_LIMIT = 4
        #         CHAR_LIMIT = 38

        #         if num_events>=1:
        #             #Insert a new line for entries which exceed the character limit
        #             for event in events:
        #                 if len(event)>CHAR_LIMIT:
        #                     split_event = twp.fill(event,CHAR_LIMIT).split("\n")
        #                     for item in split_event:
        #                         events_char_limit.append(item)
        #                 else:
        #                     events_char_limit.append(event)   
                    
        #             # Display only ROW_LIMIT number of rows. 
        #             # If the row limit is exceeded insert a row indicating how many hidden rows exist
        #             num_event_rows = len(events_char_limit)
        #             formatted_str = "\n".join(events_char_limit[0:ROW_LIMIT])
        #             if num_event_rows > ROW_LIMIT:
        #                 hidden_events = num_event_rows - ROW_LIMIT
        #                 formatted_str = formatted_str + f"\n...{hidden_events} more"
        #             table1_data.append(["Events", formatted_str])

        #     table_1_colour = np.empty_like(table1_data, dtype='object')
        #     for i, _ in enumerate(table_1_colour):
        #         table_1_colour[i] = [SECONDARY_COLOUR, SECONDARY_COLOUR]
        #     ax_table_1 = fig.add_subplot(5,3,(1,1))
        #     ax_table_1.set_title('Test Parameters', style='italic', fontsize='medium', loc='left', color=AXIS_COLOUR)
        #     tab1 = ax_table_1.table(cellText=table1_data, 
        #                             cellLoc='left', 
        #                             loc='upper left', 
        #                             cellColours=table_1_colour,
        #                             colWidths = [0.30,0.68],
        #                             bbox = [0.0,0.0,1,1]
        #                         )                   

        #     #Make event text red if it exists
        #     num_table_rows = len(table1_data)
        #     for event_cell in range(init_data_rows,num_table_rows):
        #         tab1[(event_cell,1)].get_text().set(color = 'red', wrap="False")
        #     # [t.auto_set_font_size(True) for t in [tab1]]
            
        #     #re-size event row
        #     if "Event_Data" in scenario_dict:
        #         events = scenario_dict["Event_Data"]
        #         num_events = num_table_rows-init_data_rows
        #         cellDict = tab1.get_celld()
        #         n_lines = min(len(events),ROW_LIMIT)
        #         row_height = (0.1*n_lines)
        #         padding = 0.06
        #         cellDict[(init_data_rows,0)].set_height(row_height+padding)
        #         cellDict[(init_data_rows,1)].set_height(row_height+padding)


        #     tab1.auto_set_font_size(False)
        #     tab1.set_fontsize(9)
        #     #Turn off empty plots
        #     ax_table_1.axis('off')
        #     axes[0][0].axis('off')

        # else:
        #     data = [
        #             {'Test Type': "_"},
        #             {'Test No.': "_"},
        #             {'Fault Level [MVA]': "_"},
        #             {'X2R': "_"},
        #             {'SCR': "_"},
        #     ]
        #     category = "_"

        #     # vars = {
        #     #     'Phase Step': [
        #     #                     {'spec_key':'Phase_Steps_deg_sig', 'alias':'Phase Step'}
        #     #                     ],
        #     #     'Pref Steps': [
        #     #                     {'spec_key':'Ppoc_MW_sig', 'alias':'Pref Step'}
        #     #                     ],
        #     #     'Vgrid Steps': [
        #     #                     {'spec_key':'Vpoc_pu_sig', 'alias':'Vpoc Step'}
        #     #                     ],
        #     #     'Vref Steps': [
        #     #                     {'spec_key':'Vref_pu_sig', 'alias':'Vref Step'},
        #     #                    ],
        #     #     'Fgrid Steps': [
        #     #                     {'spec_key':'Fslack_Hz_sig', 'alias':'Fgrid Step'},
        #     #                    ],
        #     #     'TOV': [
        #     #                     {'spec_key':'U_Ov', 'alias':'U_Ov (pu)'},
        #     #                    ],
        #     #     'ORT': [
        #     #                     {'spec_key':'Vslack_osc_Hz_sig', 'alias':'Osc Frequency (Hz)'},
        #     #                     {'spec_key':'Vslack_osc_phase_deg_sig', 'alias':'Osc Phase Shift (°)'},
        #     #                     {'spec_key':'Vslack_osc_amplitude_sig', 'alias':'Osc Amplitude'}
        #     #                    ],
        #     #     'PLR': [
        #     #                     {'spec_key':'PLR_Timing_Signal_sig', 'alias':'LR Timing'},
        #     #                    ],
        #     # }
        #     # if category == 'Unbalanced Faults' or category == 'Balanced Faults':
        #     #     vars['Unbalanced Faults'] = [{'spec_key':'Ures_sig', 'alias':'Vpoc Residual'}] if not pd.isna(scenario_dict['Ures_sig']) else [{'spec_key':'Zf2Zs_sig', 'alias':'Zf2Zs'},{'spec_key':'Rf_Offset_sig', 'alias':'Rf Offset'},{'spec_key':'Xf_Offset_sig', 'alias':'Xf Offset'}]
        #     #     vars['Balanced Faults'] = [{'spec_key':'Ures_sig', 'alias':'Vpoc Residual'}] if not pd.isna(scenario_dict['Ures_sig']) else [{'spec_key':'Zf2Zs_sig', 'alias':'Zf2Zs'},{'spec_key':'Rf_Offset_sig', 'alias':'Rf Offset'},{'spec_key':'Xf_Offset_sig', 'alias':'Xf Offset'}]

        #     # if category in vars:
        #     #     data.extend({dict['alias']: scenario_dict[dict['spec_key']]} for dict in vars[category])

        #     ax = fig.add_subplot(5,3,(1,1))
        #     rows = len(data)
        #     cols= 2
    
        #     for index,row in enumerate(data):
        #         row_val = list(row.values())[0]
        #         if type(row_val)==str:
        #             if len(row_val)<=35 and self.style_class.tables.data_fontsize <=8:
        #                 scaled_fontsize = self.style_class.tables.data_fontsize

        #             elif len(row_val)<=42 and self.style_class.tables.data_fontsize >7:
        #                 scaled_fontsize = 7

        #             elif len(row_val)<=48 and self.style_class.tables.data_fontsize >6:
        #                 scaled_fontsize = 6

        #             elif len(row_val)<=63 and self.style_class.tables.data_fontsize >5:
        #                 scaled_fontsize = 5

        #             elif len(row_val)<=73 and self.style_class.tables.data_fontsize >4:
        #                 scaled_fontsize = 4

        #             elif len(row_val)<=98 and self.style_class.tables.data_fontsize >3:
        #                 scaled_fontsize = 3

        #             elif len(row_val)<=112 and self.style_class.tables.data_fontsize >2.90:
        #                 scaled_fontsize = 2.90

        #             elif len(row_val)<=116 and self.style_class.tables.data_fontsize >2.75:
        #                 scaled_fontsize = 2.75

        #             elif len(row_val)<=127 and self.style_class.tables.data_fontsize >2.5:
        #                 scaled_fontsize = 2.5

        #             elif len(row_val)<=139 and self.style_class.tables.data_fontsize >2.25:
        #                 scaled_fontsize = 2.25

        #             elif len(row_val)<=147 and self.style_class.tables.data_fontsize >2:
        #                 scaled_fontsize = 2

        #             else:
        #                 scaled_fontsize = self.style_class.tables.data_fontsize
        #         else:
        #             scaled_fontsize = self.style_class.tables.data_fontsize

        #         ax.text(x=0, y=len(data)-index-1, s=list(row.keys())[0], va='center', ha='left', weight='bold', fontsize=self.style_class.tables.header_fontsize)
        #         ax.text(x=1.05, y=len(data)-index-1, s=list(row.values())[0], va='center', ha='left', fontsize=scaled_fontsize)


        #     for row in range(rows):
        #         #row gridlines
        #         ax.plot(
        #             [0, cols + 1],
        #             [row -.5, row - .5],
        #             ls=':',
        #             lw='.5',
        #             c=self.style_class.tables.gridline_colour
        #         )
        #     #column highlight
        #     rect = patches.Rectangle(
        #         (0, -.5),  # bottom left starting position (x,y)
        #         1.0,  # width
        #         len(data),  # height
        #         ec='none',
        #         fc=self.style_class.tables.shading_colour,
        #         alpha=.2,
        #         zorder=-1
        #     )
        #     #header divider
        #     ax.plot([1, 1], [-.5, rows-0.5], lw='.5', c='black')

        #     ax.add_patch(rect)
        #     ax.axis('off')
        #     axes[0][0].axis('off')
        #     ax.set_title(
        #         "Filename",
        #         loc='left',
        #         fontsize=12,
        #         weight='bold'
        #     )


        # Adjust layout
        plt.tight_layout()
        plt.savefig(png_path, bbox_inches="tight",edgecolor=fig.get_edgecolor())
        plt.savefig(pdf_path, bbox_inches="tight", format="pdf",edgecolor=fig.get_edgecolor())
        plt.clf()
        plt.close()
