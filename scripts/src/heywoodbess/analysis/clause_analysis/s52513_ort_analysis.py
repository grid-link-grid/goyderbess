import os
import json
import pandas as pd
import numpy as np
from typing import List, Callable, Optional
from tqdm.auto import tqdm
from pallet.pscad.Psout import Psout
from pallet.dsl.ParsedDslSignal import ParsedDslSignal
from dataclasses import dataclass
import matplotlib.pyplot as plt
import math


@dataclass
class SubplotData:
    osc_freq_hz : float
    vpoc_pu :  pd.Series
    qpoc_mvar : float


def produce_ort_plots(
        psout_paths : List[str],
        vslack_osc_freq_spec_key : str,
        vslack_osc_amp_spec_key : str,
        vpoc_channel_key : str,
        qpoc_channel_key : str,
        output_path : str,
        periods_per_subplot : float 
    ):

    def make_suplot(
          subplot_data: SubplotData,
          ax: plt.Axes,
          nperiods: float,
          start_time_s: float,
        ):

        period_s = 1/subplot_data.osc_freq_hz
        end_time_s = start_time_s+(period_s*nperiods)

        osc_duration_displayed = end_time_s-start_time_s
        start_time_s = max(0,start_time_s-0.1*osc_duration_displayed)

        series_vpoc_pu = subplot_data.vpoc_pu[(subplot_data.vpoc_pu.index > start_time_s)&(subplot_data.vpoc_pu.index < end_time_s)]
        series_qpoc_mvar = subplot_data.qpoc_mvar[(subplot_data.qpoc_mvar.index > start_time_s)&(subplot_data.qpoc_mvar.index < end_time_s)]

        ax.plot(series_vpoc_pu.index,series_vpoc_pu, color = "blue")
        ax.twinx().plot(series_qpoc_mvar.index,series_qpoc_mvar, color = "red",linestyle='--')
        ax.set_title(f"{str(subplot_data.osc_freq_hz)} hz", loc='left',pad=-0.5)
        ax.grid(True)

    cols = 3
    rows = int(math.ceil(len(psout_paths)/cols))
    fig, axes = plt.subplots(nrows=rows, ncols=cols, figsize=(16, 9), sharex=False)
    cm = 1 / 2.54
    cm_per_col = 18
    cm_per_row = 8
    fig.set_size_inches(cm_per_col * cols * cm, cm_per_row * rows * cm)  
    plt.subplots_adjust(left=0.15, right=0.85, bottom=0.1, top=0.85, wspace=0.25, hspace=0.34)


    axes = axes.flatten()
    for i in tqdm(range(len(psout_paths)), desc="Preparing ORT Summary Plot"):

        psout_path = psout_paths[i]
        json_path = psout_path.split(".psout")[0] + ".json" 

        with open(json_path, 'r') as f:
            spec = json.load(f)    

        for dir in [os.path.dirname(output_path)]:
            if not os.path.exists(dir):
                os.makedirs(dir)

        osc_cmd = ParsedDslSignal.parse(spec[vslack_osc_amp_spec_key]).signal
        osc_init_time_s = osc_cmd.pieces[0].time_start

        init_time_sec = float(spec["substitutions"]["TIME_Full_Init_Time_sec"])
        df = Psout(psout_path).to_df(secs_to_remove=init_time_sec)
        osc_freq = spec[vslack_osc_freq_spec_key]

        subplot_data = SubplotData(
                osc_freq_hz=osc_freq,
                vpoc_pu=df[vpoc_channel_key],
                qpoc_mvar=df[qpoc_channel_key]
         )

        make_suplot(
            subplot_data = subplot_data,
            ax = axes[i],
            nperiods = periods_per_subplot,
            start_time_s = osc_init_time_s
        )

    plt.savefig(output_path, bbox_inches="tight", format="png")
    plt.clf()
    plt.close()