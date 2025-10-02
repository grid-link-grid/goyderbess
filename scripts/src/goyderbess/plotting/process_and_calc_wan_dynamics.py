import sys

from pathlib import Path
PSSBIN_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSBIN"""
PSSPY_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSPY39"""
sys.path.append(PSSBIN_PATH)
sys.path.append(PSSPY_PATH)
import psse34
import psspy

import os
import pandas as pd
from icecream import ic
import json
import numpy as np
import inspect

from pallet.specs import load_specs_from_xlsx
from pallet.psse.PssePallet import PssePallet
from pallet.psse import system, statics
from pallet.psse.vslack_calculator import calc_vslack
from pallet.psse.case.Bus import Bus, BusData
from pallet.psse.case.Machine import Machine, MachineData

from gridlink.utils.wan.wan_utils import find_files

from pallet import now_str
from pallet.psse.Out import Out

from tqdm.auto import tqdm
from typing import List, Optional, Callable



INV_LVRT_OUT_PU = 0.851
INV_LVRT_IN_PU = 0.85
INV_HVRT_OUT_PU = 1.179
INV_HVRT_IN_PU = 1.18

PPC_HVRT_PU = 1.25
PPC_LVRT_PU = 0.8


def calc_iq(
        vbase_kV: float,
        sbase_mva: float,
        v_pu: pd.DataFrame,
        p_mw: pd.DataFrame,
        q_mvar: pd.DataFrame
    ):
    
    Id_kA = p_mw/(np.sqrt(3)*v_pu*vbase_kV)
    Iq_kA = q_mvar/(np.sqrt(3)*v_pu*vbase_kV)


    Is_kA = sbase_mva/(np.sqrt(3)*vbase_kV)
    Iq_PU = Iq_kA/Is_kA
    Id_PU = Id_kA/Is_kA

    return Iq_PU, Id_PU



def pre_process_dataframe(df:pd.DataFrame):

    df_copy = df.copy()


    df_copy["NOTE_INV_LVRT_Out_pu"] = INV_LVRT_OUT_PU
    df_copy["NOTE_INV_LVRT_In_pu"] = INV_LVRT_IN_PU

    df_copy["NOTE_INV_HVRT_Out_pu"] = INV_HVRT_OUT_PU
    df_copy["NOTE_INV_HVRT_In_pu"] = INV_HVRT_IN_PU

    df_copy["NOTE_PPC_HVRT_pu"] = PPC_HVRT_PU
    df_copy["NOTE_PPC_LVRT_pu"] = PPC_LVRT_PU

    df_copy["Hz_POI"] = df_copy["POC_F_PU"]*50.0 + 50.0

    df_copy['Main_Tx_HVMV_Ratio'] = df_copy['V_POC_PU']/df_copy['V_MV_PU']


    POC_VBASE_kV = 275.0
    POC_SBASE_MVA = 423.2
    POC_SBASE_MAX_CONT_CURRENT_MVA = 306.4280187
    
    Iq_PU, Id_PU = calc_iq(
        vbase_kV=POC_VBASE_kV,
        sbase_mva=POC_SBASE_MVA,
        v_pu=df_copy["V_POC_PU"],
        p_mw=df_copy["PPOC_MW"],
        q_mvar=df_copy["QPOC_MVAR"]
    )

    df_copy["POC_Iq_PU"] = Iq_PU
    df_copy["POC_Id_PU"] = Id_PU

    Iq_PU_max_cont, Id_PU_max_cont = calc_iq(
        vbase_kV=POC_VBASE_kV,
        sbase_mva=POC_SBASE_MAX_CONT_CURRENT_MVA,
        v_pu=df_copy["V_POC_PU"],
        p_mw=df_copy["PPOC_MW"],
        q_mvar=df_copy["QPOC_MVAR"]
    )

    df_copy["POC_Iq_PU_max_cont"] = Iq_PU_max_cont
    df_copy["POC_Id_PU_max_cont"] = Id_PU_max_cont

    INV1_VBASE_kV = 0.69


    INV1_SBASE_MVA = 105.8

    PSSE_SYSTEM_MBASE_MVA = 100.0

    INV1_Iq_PU, INV1_Id_PU = calc_iq(
        vbase_kV=INV1_VBASE_kV,
        sbase_mva=INV1_SBASE_MVA,
        v_pu=df_copy["V_HYBESS_PU"],
        p_mw=df_copy["PMEAS_INV_PU"]*PSSE_SYSTEM_MBASE_MVA,
        q_mvar=df_copy["QMEAS_INV_PU"]*PSSE_SYSTEM_MBASE_MVA
    )

    df_copy["INV1_Iq_PU"] = INV1_Iq_PU 
    df_copy["INV1_Id_PU"] = INV1_Id_PU 


    df_copy["Irms_HV_pu_max_cont"] = np.sqrt(Iq_PU_max_cont**2 + Id_PU_max_cont**2) 
    df_copy["Irms_HV_pu"] = np.sqrt(Iq_PU**2 + Id_PU**2) 
    df_copy["INV1_Irms_LV_pu"] = np.sqrt(INV1_Iq_PU**2 + INV1_Id_PU**2)  


    # TO BE EVALUATED
    df_copy["poc_p_err_MW"] = 0.0 
    df_copy["poc_q_err_MVAr"] = 0.0
    df_copy["poc_v_err_adj_pu"] = 0.0


    # Calculate the HV/MV ratio for the main transformer. Insert a column into the dataframe that contains that information
    # Voltage signal on primary side of Tx / Voltage signal on secondary side of Tx
    # UPDATE THESE VALUES WHEN WE HAVE THE ACTUAL SIGNALS FOR goyderbess: Irms_HV_pu and VRMS_POC_PU ###########################################################
    ############################################################################################################################
    
    df_copy['Main_Tx_HVMV_Ratio'] = df_copy['Irms_HV_pu']/df_copy['V_POC_PU']
    # df_copy = df_copy.assign(Main_Tx_HVMV_Ratio)

    # TO BE EVALUATED
    df_copy["poc_p_err_MW"] = 0.0 
    df_copy["poc_q_err_MVAr"] = 0.0
    df_copy["poc_v_err_adj_pu"] = 0.0

    return df_copy
