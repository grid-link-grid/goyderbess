import os
import sys
import pandas as pd
import shutil

from pallet.specs import load_specs_from_xlsx

SPEC_PATH = os.path.join(os.getcwd(), "GESF_Specs_Luke.xlsx")
OUTPUT_DIR = os.path.join(os.getcwd(), "Output")
COLMAP_PATH = os.path.join(os.getcwd(), "column_mapping.csv")

SHEETS_TO_PROCESS = [
    "5251",
    "5253",
    "5254_CUO",
    "5254_Withstand",
    "52511_Fgrid",
    # "52513_Vgrid",
    "52513_Vref",
    "52513_Qref",
    "52513_PFref",
    "5257",
#----------------DMAT--------------------------------------------

    "3215_ORT", #PSCAD ONLY
    "3210_3214_Vgrid",
    "3211_Pref",
    "3210_PFref",
    "3210_Vref",
    "3210_Qref",
    "3212_Fgrid",
    "324_325_Faults",
    "329_TOV",
    "3220_Input_Power",
    "3216_PhaseSteps",
    # "3218_3219_SCR_Change_Faults"
]


def manipulations(input_df : pd.DataFrame) -> pd.DataFrame:
    df = input_df.copy()

    # Fault imepdance
    df["Fault_Impedance"] = None
    for idx, row in df.iterrows():
        if pd.notna(row["Ures_sig"]):
            df.at[idx, "Fault_Impedance"] = f"{row['Ures_sig']} pu"
        elif pd.notna(row["Zf2Zs_sig"]):
            if pd.notna(row["Rf_Offset_sig"]) and row["Rf_Offset_sig"] > 0:
                df.at[idx, "Fault_Impedance"] = f"{row['Rf_Offset_sig']} Ω"
            else:
                df.at[idx, "Fault_Impedance"] = f"Zf = {row['Zf2Zs_sig']} Zs"
    


    return df




if __name__ == "__main__":
    pd.reset_option('all')

    if os.path.exists(OUTPUT_DIR) and os.path.isdir(OUTPUT_DIR):
        shutil.rmtree(OUTPUT_DIR)
    os.mkdir(OUTPUT_DIR)

    colmap_df = pd.read_csv(COLMAP_PATH)
    print(colmap_df)
    spec_df = load_specs_from_xlsx(SPEC_PATH, sheet_names=SHEETS_TO_PROCESS, add_sheet_name_column=True)
    spec_df = manipulations(spec_df)

    spec_df["Sheet_Name"] = spec_df["Sheet_Name"].astype(str)

    # print(spec_df)


    for _, sheet_colmap in colmap_df.iterrows():
        sheet_colmap = sheet_colmap[sheet_colmap.notna()]
        sheet_name = str(sheet_colmap["Sheet_To_Update"])
        sheet_colmap = sheet_colmap.drop(labels=["Sheet_To_Update"])
        # print(sheet_colmap)

        sheet_df = spec_df[spec_df["Sheet_Name"] == sheet_name]
        # print(sheet_df)
        renaming_dict = dict(zip([str(x) for x in sheet_colmap.index.tolist()], [str(x) for x in sheet_colmap.values.tolist()]))
        sheet_df = sheet_df.rename(columns=renaming_dict)
        sheet_df = sheet_df[[str(x) for x in sheet_colmap.values.tolist()]]
        sheet_df = sheet_df.apply(lambda x: x.round(2) if pd.api.types.is_numeric_dtype(x) else x)
        # print(sheet_df)



        # Reference and results columns
        # sheet_df["Figure"] = [f"Figure {i+1}" for i in range(len(sheet_df))]
        sheet_df["Results"] = "Acceptable"

        # print(sheet_df)
        sheet_name = sheet_name.replace("_", "-")
        sheet_df.to_csv(os.path.join(OUTPUT_DIR, f"{sheet_name}-test-table.csv"), index=False)

        
