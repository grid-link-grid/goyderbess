import os
import json
import pandas as pd
import math
from pallet import load_specs_from_csv, load_specs_from_xlsx
from pallet.pscad.PscadPallet import PscadPallet
from pallet.dsl.ParsedDslSignal import ParsedDslSignal



#i will need to put this in the top of any script that needs to call this script:
#from cgbess.utils.vslack import calc_vslacks_with_caching


def calc_vslacks_with_caching(spec, use_cache : bool, spec_path, vbase_kv, fbase_hz):
    '''This function calculates Vslack for the studies outlined in the spec.

    use_cache = False will delete any cached vslack values, recalculate all vslacks, and cache them in vslack_cache.csv
    
    use_cache = True will check if a vslack value is present in vslack_cache.csv for the parameters in every row of the spec.
        
    ----If a vslack value is present, it will put it in the spec dataframe that this function outputs.
        
    ----If a vslack value is not present for a given row in the spec, the function will calculate a value and put it in the spec dataframe that this function outputs, as well as putting it in the cache
    
    Pass the path to the spec (XSLX_PATH) to spec_path, and this function will save the vslack_cache in the same folder'''
    
    cache_path = os.path.join(os.path.dirname(spec_path),"vslack_cache.csv")
    print(f'Cache is located at : {cache_path}')

    p = PscadPallet()
    spec["constant_vslack_pu_sig"] = 999
    if 'Calc_TOV' not in spec.columns:
        spec['Calc_TOV'] = False
        spec['U_Ov'] = 0.0
    spec['Calc_TOV'] = spec['Calc_TOV'].fillna(False)
    p.set_spec(spec)
    p.set_initialisation_column_names(
        fault_level_mva_col_name="Grid_FL_MVA_sig",
        x2r_col_name="Grid_X2R_sig",
        vpoc_pu_col_name="Vpoc_pu_sig",
        ppoc_mw_col_name="Ppoc_MW_sig",
        qpoc_mvar_col_name="Qpoc_MVAr_init",
        vslack_pu_col_name="constant_vslack_pu_sig",
        tov_voltage_pu_col_name="U_Ov",
        tov_shunt_mvar_col_name="TOV_MVAr",
        vpoc_disturbance_pu_col_name="Vpoc_disturbance_pu_sig",
    )
    p.set_bases(vbase_kv=vbase_kv, fbase_hz=fbase_hz)
    
    def use_cache_false(palletObject):
        p = palletObject
        if os.path.exists(cache_path):
            os.remove(cache_path)

        p.add_vslacks_to_spec(verbose=False)
        p.add_tov_shunt_size_to_spec(filter_column="Calc_TOV")
        p.spec["TOV_Shunt_C_uF_sig"] = (p.spec["TOV_MVAr"]* 1E6/(2 * math.pi * 50 * (1.0 * p.vbase_kv)**2))

        spec = p.spec.copy()
        print(spec)

        # Step 1: Replace empty values with constant_vslack_pu_sig
        spec.loc[~spec['Vslack_pu_sig'].astype(str).str.contains(r"\$VSLACK", na=False), 'Vslack_pu_sig'] = spec["constant_vslack_pu_sig"]

        # Step 2: Ensure all values are strings before applying .replace()
        spec['Vslack_pu_sig'] = spec['Vslack_pu_sig'].astype(str)

        # Step 3: Replace "$VSLACK" with Vslack_pu_sig values
        spec['Vslack_pu_sig'] = spec.apply(
            lambda row: row['Vslack_pu_sig'].replace("$VSLACK", str(row['constant_vslack_pu_sig'])), axis=1
            )
        
        spec_for_vslack = spec

        # mask = spec_for_vslack['Ppoc_MW_sig'].apply(lambda x: isinstance(x, str))
        # spec_for_vslack.loc[mask, 'Ppoc_MW_sig'] = spec_for_vslack.loc[mask, 'Ppoc_MW_sig'].apply(
        # lambda val: ParsedDslSignal.parse(val).signal.initial_value)

        # mask = spec_for_vslack['Qpoc_MVAr_sig'].apply(lambda x: isinstance(x, str))
        # spec_for_vslack.loc[mask, 'Qpoc_MVAr_sig'] = spec_for_vslack.loc[mask, 'Qpoc_MVAr_sig'].apply(
        # lambda val: ParsedDslSignal.parse(val).signal.initial_value)

        vslack_cache = spec_for_vslack.filter(["Grid_FL_MVA_sig","Grid_X2R_sig","Vpoc_pu_sig","Ppoc_pu","Qpoc_pu","Vslack_pu_sig", "constant_vslack_pu_sig", "U_Ov", "TOV_MVAr", "TOV_Shunt_C_uF_sig"], axis=1)
        vslack_cache['Vslack_pu_sig'] = vslack_cache['constant_vslack_pu_sig']
        vslack_cache.drop(columns = 'constant_vslack_pu_sig')
        vslack_cache = vslack_cache.drop_duplicates()
        vslack_cache.to_csv(cache_path, index = False)

        return spec

    if use_cache == False:
        spec = use_cache_false(p)
    
    if use_cache == True:
        if os.path.exists(cache_path):
            vslack_cache = pd.read_csv(cache_path)
        else:
            print('------------------------------------------------------------------')
            print('')
            print('vslack_cache.csv does not exist. Now running as if use_cache=False')
            print('')
            print('------------------------------------------------------------------')
            spec = use_cache_false(p)
            return spec
            

        # check each row in the spec against the calc_vslack column and if true, check for identical rows in the cache and take the vslack value into the spec. if calc_vslack is false, calculate the vslack
        spec['TOV_MVAr'] = 0.0
        spec['TOV_Shunt_C_uF_sig'] = 0.0
        

        cols_to_match = ["Grid_FL_MVA_sig","Grid_X2R_sig","Vpoc_pu_sig","Ppoc_pu","Qpoc_pu", "U_Ov"]


        spec_copy = spec
        # spec_copy['Ppoc_MW_sig_copy'] = spec_copy['Ppoc_MW_sig']
        # spec_copy['Qpoc_MVAr_sig_copy'] = spec_copy['Qpoc_MVAr_sig']
        # vslack_cache['Ppoc_MW_sig_copy'] = vslack_cache['Ppoc_MW_sig']
        # vslack_cache['Qpoc_MVAr_sig_copy'] = vslack_cache['Qpoc_MVAr_sig']
        
        # mask = spec_copy['Ppoc_MW_sig_copy'].apply(lambda x: isinstance(x, str))
        # spec_copy.loc[mask, 'Ppoc_MW_sig_copy'] = spec_copy.loc[mask, 'Ppoc_MW_sig_copy'].apply(
        # lambda val: ParsedDslSignal.parse(val).signal.initial_value)

        # mask = spec_copy['Qpoc_MVAr_sig_copy'].apply(lambda x: isinstance(x, str))
        # spec_copy.loc[mask, 'Qpoc_MVAr_sig_copy'] = spec_copy.loc[mask, 'Qpoc_MVAr_sig_copy'].apply(
        # lambda val: ParsedDslSignal.parse(val).signal.initial_value)


        if 'U_Ov' not in spec.columns:
            spec['U_Ov'] = 0.0

        merged = pd.merge(
            spec_copy,
            vslack_cache[cols_to_match+["Vslack_pu_sig", "TOV_MVAr", "TOV_Shunt_C_uF_sig"]],  # ensure only relevant columns + target column
            on=cols_to_match,
            how='left',
            suffixes=('', '_from_cache')
        )

        print(merged['Vslack_pu_sig_from_cache'])

        merged["constant_vslack_pu_sig"] = merged["Vslack_pu_sig_from_cache"]
        merged["TOV_MVAr"] = merged["TOV_MVAr_from_cache"]
        merged["TOV_Shunt_C_uF_sig"] = merged["TOV_Shunt_C_uF_sig_from_cache"]
        print(merged["TOV_Shunt_C_uF_sig"])

        spec_nan = merged[merged["constant_vslack_pu_sig"].isna()]
        merged_cleaned = merged.drop(spec_nan.index)

        if len(spec_nan) > 0:
            # print(spec_nan)
            if 'Calc_TOV' not in spec.columns:
                spec['Calc_TOV'] = False
            spec['Calc_TOV'] = spec['Calc_TOV'].fillna(False)
            p.set_spec(spec_nan)
            p.add_vslacks_to_spec(verbose=False)

            p.add_tov_shunt_size_to_spec(filter_column ="Calc_TOV")
            p.spec["TOV_Shunt_C_uF_sig"] = (p.spec["TOV_MVAr"]* 1E6/(2 * math.pi * 50 * (1.0 * p.vbase_kv)**2))

            spec_nan = p.spec.copy()
            
            spec_prelim = pd.concat([merged_cleaned,spec_nan], ignore_index=True)
            print(spec_prelim["constant_vslack_pu_sig"])
            print(spec_prelim["Vslack_pu_sig"])

            # Step 1: Replace empty values with constant_vslack_pu_sig
            spec_prelim.loc[~spec_prelim['Vslack_pu_sig'].astype(str).str.contains(r"\$VSLACK", na=False), 'Vslack_pu_sig'] = spec_prelim["constant_vslack_pu_sig"]

            # Step 2: Ensure all values are strings before applying .replace()
            spec_prelim['Vslack_pu_sig'] = spec_prelim['Vslack_pu_sig'].astype(str)

            # Step 3: Replace "$VSLACK" with Vslack_pu_sig values
            spec_prelim['Vslack_pu_sig'] = spec_prelim.apply(
                lambda row: row['Vslack_pu_sig'].replace("$VSLACK", str(row['constant_vslack_pu_sig'])), axis=1
                )
        else:
            spec_prelim = merged
            spec_prelim.loc[~spec_prelim['Vslack_pu_sig'].astype(str).str.contains(r"\$VSLACK", na=False), 'Vslack_pu_sig'] = spec_prelim["constant_vslack_pu_sig"]

            # Step 2: Ensure all values are strings before applying .replace()
            spec_prelim['Vslack_pu_sig'] = spec_prelim['Vslack_pu_sig'].astype(str)

            # Step 3: Replace "$VSLACK" with Vslack_pu_sig values
            spec_prelim['Vslack_pu_sig'] = spec_prelim.apply(
                lambda row: row['Vslack_pu_sig'].replace("$VSLACK", str(row['constant_vslack_pu_sig'])), axis=1
                )

        spec = spec_prelim

        vslack_cache = pd.concat([vslack_cache, spec.filter(["Grid_FL_MVA_sig","Grid_X2R_sig","Vpoc_pu_sig","Ppoc_pu","Qpoc_pu",'constant_vslack_pu_sig', "U_Ov", "TOV_MVAr", "TOV_Shunt_C_uF_sig"])])
        vslack_cache['Vslack_pu_sig'] = vslack_cache['constant_vslack_pu_sig']
        vslack_cache.reset_index(inplace=True,drop=True)
        # vslack_cache = spec.filter(["Grid_FL_MVA_sig","Grid_X2R_sig","Vpoc_pu_sig","Ppoc_MW_sig","Qpoc_MVAr_sig","Vslack_pu_sig", "U_Ov", "TOV_MVAr", "TOV_Shunt_C_uF_sig"], axis=1)
        vslack_cache = vslack_cache.drop_duplicates()
                



        vslack_cache.to_csv(cache_path, index = False)
    
    print('end')

    return spec
