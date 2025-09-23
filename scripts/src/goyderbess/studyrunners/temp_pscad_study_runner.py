import sys
import os
import json
#os.environ["POWERKIT_LICENSE_KEY"] = "396119-1D6701-6F484D-30E154-E30B6A-V3"
import pandas as pd 
import numpy as np
import math
import glob

import ast
from pallet.pscad.PscadPallet import PscadPallet
from pallet.pscad.launching import launch_pscad
from pallet.pscad.validation import ValidatorBehaviour
from pallet import load_specs_from_csv, load_specs_from_xlsx
from pallet import now_str
from heywoodbess.plotting.heywoodbessPscadPlotter import heywoodbessPscadPlotter
from pathlib import Path
from heywoodbess.utils.vslack import calc_vslacks_with_caching

def run_pscad_studies_temp(
        SKIP_PANDAPOWER: bool,
        SHEETS_TO_PROCESS,
        volley_size,
        plotter: object,
        XLSX_PATH,
        MODEL_DIR,
        TEMP_DIR,
        RESULTS_DIR
                    ):


    if SKIP_PANDAPOWER:
        spec = load_specs_from_csv("skip_spec.csv")
    else:
        spec = load_specs_from_xlsx(XLSX_PATH, sheet_names=SHEETS_TO_PROCESS)
        spec = spec[spec["PSCAD"] == True]
        spec["Vslack_pu_sig"] = None
        spec["TOV_MVAr"] = None

        spec["Calc_Vslack"] = True
        spec["Calc_TOV"] = spec["Category"] == "TOV"

    # spec = spec[spec["Test No"] == "155"]

    print("********* SPEC ****************")
    print(spec)
    print("******************************")

    p = PscadPallet()
    p.set_spec(spec)
    p.set_initialisation_column_names(
        fault_level_mva_col_name="Grid_FL_MVA_sig",
        x2r_col_name="Grid_X2R_sig",
        vpoc_pu_col_name="Vpoc_pu_sig",
        ppoc_mw_col_name="Ppoc_MW_sig",
        qpoc_mvar_col_name="Qpoc_MVAr_sig",
        vslack_pu_col_name="Vslack_pu_sig",
        # tov_voltage_pu_col_name="U_Ov",
        # tov_shunt_mvar_col_name="TOV_MVAr",
        vpoc_disturbance_pu_col_name="Vpoc_disturbance_pu_sig",
    )
    p.set_bases(vbase_kv=33.0, fbase_hz=50.0)

    if not SKIP_PANDAPOWER:
        p.add_vslacks_to_spec(filter_column="Calc_Vslack", verbose=False)
        # p.add_tov_shunt_size_to_spec(filter_column="Calc_TOV")
        p.spec.to_csv("skip_spec.csv")

    # Validate here if needed

    
    pscad, project = launch_pscad(
        workspace_path=os.path.join(MODEL_DIR, "Workspace.pswx"),
        project_name="HyCon_Demo",
        copy_to_dir=TEMP_DIR,
    )

    gs_backup_path = os.path.join(TEMP_DIR, "last_known_good_substitutions.json")
    p.backup_global_substitutions(
        project=project,
        json_path=gs_backup_path,
    )
    p.validate_global_substitutions(
        project=project,
        json_path=gs_backup_path,
        behaviour=ValidatorBehaviour.WARN,
    )
    
    p.run(
        pscad=pscad,
        project=project,
        testbench_map_path=os.path.join(MODEL_DIR, "testbench_mapping.json"),
        volley_size=volley_size,
        quit_after=False,
        results_dir=RESULTS_DIR,
        subfolder_column="Category",
        filename_column="File_Name",
        plotter=plotter,
        trace_all=True,
        prioritised_column="Category",
        # prioritised_values=["Fgrid Steps"],
    )