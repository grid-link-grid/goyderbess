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
from goyderbess.plotting.goyderbessPscadPlotter import goyderbessPscadPlotter
from pathlib import Path
from goyderbess.utils.vslack import calc_vslacks_with_caching
from pallet.pandapower.initialisation import calc_tov_shunt_mvar_from_ppoc_qpoc_vpoc_vslack

VBASE_KV = 275
FBASE_HZ = 50


def run_pscad_studies(
            USE_VSLACK_CACHE: bool,
            spec,
            project_name:str,
            volley_size,
            plotter: object,
            XLSX_DIR,
            MODEL_DIR,
            TEMP_DIR,
            RESULTS_DIR):

    spec = calc_vslacks_with_caching(spec=spec, use_cache=USE_VSLACK_CACHE, spec_path=XLSX_DIR, vbase_kv=VBASE_KV, fbase_hz=FBASE_HZ)
   
    

    print("********* SPEC ****************")
    print(spec)
    print("******************************")

    p = PscadPallet()
    p.set_spec(spec)

    p.set_bases(vbase_kv=VBASE_KV, fbase_hz=FBASE_HZ)

    # CALC_TOV_SHUNT_VAR = True
    # if CALC_TOV_SHUNT_VAR:

    #     p.set_initialisation_column_names(
    #         fault_level_mva_col_name="Grid_FL_MVA_sig",
    #         x2r_col_name="Grid_X2R_sig",
    #         vpoc_pu_col_name="Vpoc_pu_sig",
    #         ppoc_mw_col_name="Ppoc_MW_sig",
    #         qpoc_mvar_col_name="Qpoc_MVAr_sig",
    #         vslack_pu_col_name="Vslack_pu_sig",
    #         tov_voltage_pu_col_name="U_Ov",
    #         tov_shunt_mvar_col_name="TOV_MVAr",
    #         vpoc_disturbance_pu_col_name="Vpoc_disturbance_pu_sig",
    #     )
        # p.add_tov_shunt_size_to_spec(filter_column="Calc_TOV")
        # print(p.spec["TOV_MVAr"])
        # p.spec["TOV_Shunt_C_uF_sig"] = (p.spec["TOV_MVAr"]* 1E6/(2 * math.pi * 50 * (1.0 * p.vbase_kv)**2))
        # tov_spec.to_csv("spec_tov_values.csv")
        # print("Finished calculating TOV. TOV shunt mvar values stored in tov_shunt_mvar.csv -->  Aborting..")
        # spec_copy = spec.copy()
        # spec_copy = spec_copy[spec_copy["Category"].isin(["TOV","CSR TOV"])]
        # tov_shunt_mvar = []
        # if spec_copy.empty:
        #     for row in spec_copy.iterrows():
                
        #         tov_shunt_mvar_value = calc_tov_shunt_mvar_from_ppoc_qpoc_vpoc_vslack(
        #             fault_level_mva=row["Grid_FL_MVA_sig"],
        #             x2r=row["Grid_X2R_sig"],
        #             vbase_kv=VBASE_KV,
        #             ppoc_mw=row["Ppoc_MW_sig"],
        #             qpoc_mvar=row["Qpoc_MVAr_sig"],
        #             post_shunt_vpoc_pu=row["U_Ov"],
        #             vslack_pu=row["Vslack_pu_sig"],
        #             initial_step_size_pu=0.01,
        #             convergence_tolerance_pu=0.0001,
        #             max_iterations=100,
        #             verbose=False)
        #         tov_shunt_mvar.append(tov_shunt_mvar_value)
        #     columns = ["Deliverable","Category","Grid_FL_MVA_sig","Grid_X2R_sig","Ppoc_MW_sig","Qpoc_MVAr_sig","U_Ov","Vslack_pu_sig"] 
        #     tov_spec = spec[columns]
        #     tov_spec["TOV_MVAr"] = tov_shunt_mvar
        #     tov_spec.to_csv("spec_tov_values.csv")
        #     print("Finished calculating TOV. TOV shunt mvar values stored in tov_shunt_mvar.csv -->  Aborting..")
        #     sys.exit()


    p.spec.to_csv("skip_spec.csv")
    # Validate here if needed
    
    pscad, project = launch_pscad(
        workspace_path=os.path.join(MODEL_DIR, "Workspace.pswx"),
        project_name=project_name,
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
        plotters=[plotter],
        trace_all=False,
        prioritised_column="Category",
        # prioritised_values=["Fgrid Steps"],
        save_output_as_pickle=False,
        save_output_as_psout=True,
    )

    return