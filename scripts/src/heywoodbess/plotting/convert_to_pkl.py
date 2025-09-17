
import os
import glob
import pandas as pd
import numpy as np
import json
import sys
import pickle



def psout_to_pkl(x86,PSCAD_INPUT_DIR):
    if not x86:
        from pallet.pscad.Psout import Psout
    folders = [item for item in os.listdir(PSCAD_INPUT_DIR) if os.path.isdir(os.path.join(PSCAD_INPUT_DIR, item))]
    print(folders)
    for folder in folders:
        folder_path = os.path.join(PSCAD_INPUT_DIR,folder)
        for file in os.listdir(folder_path):
            if file.endswith('.psout'):
                psout_path = os.path.join(folder_path,file)
                pkl_path = os.path.join(folder_path, file.replace(".psout",".pkl"))
                psout_df = Psout(psout_path).to_df()
                psout_df.to_pickle(pkl_path)

                os.remove(psout_path)

                print(f"Converted: {file} --> {file.replace('.psout','.pkl')} and removed the original")
    print("Done")

    return 



def test(x86,PSCAD_INPUT_DIR):
    if not x86:
        from pallet.pscad.Psout import Psout
    folders = [item for item in os.listdir(PSCAD_INPUT_DIR) if os.path.isdir(os.path.join(PSCAD_INPUT_DIR, item))]
    print(folders)
    for folder in folders:
        print(f"this is the folder: {folder}")
        folder_path = os.path.join(PSCAD_INPUT_DIR,folder)
        for file in os.listdir(folder_path):
            print(file)
            # print(file)
            if file.endswith('.psout'):
                print("ends with psout")
            #     psout_path = os.path.join(folder_path,file)
            #     pkl_path = os.path.join(folder_path, file.replace(".psout",".pkl"))
            #     psout_df = Psout(psout_path).to_df()
            #     psout_df.to_pickle(pkl_path)

            #     os.remove(psout_path)

            #     print(f"Converted: {file} --> {file.replace('.psout','.pkl')} and removed the original")
    print("Done")

    return 




def out_to_pkl(x86,PSSE_INPUT_DIR):
    print('hello')
    if x86:
        from pallet.psse.Out import Out
    folders = [item for item in os.listdir(PSSE_INPUT_DIR) if os.path.isdir(os.path.join(PSSE_INPUT_DIR, item))]
    print(folders)
    for folder in folders:
        folder_path = os.path.join(PSSE_INPUT_DIR,folder)
        for file in os.listdir(folder_path):
            if file.endswith('.out'):
                out_path = os.path.join(folder_path,file)
                pkl_path = os.path.join(folder_path, file.replace(".out",".pkl"))
                out_df = Out(out_path).to_df()
                out_df.to_pickle(pkl_path)

                os.remove(out_path)

                print(f"Converted: {file} --> {file.replace('.out','.pkl')} and removed the original")
    print("Done")

    return 


