import os
import sys



convert_psse_to_pkl = True #Set to true if you want to convert psse out files to pkls. 


# psout_to_pkl(x86=False,PSCAD_INPUT_DIR=PSCAD_INPUT_DIR)

def convert_to_pkl(convert_psse_to_pkl:bool):


    if convert_psse_to_pkl:
        from heywoodbess.plotting.convert_to_pkl import out_to_pkl
        PSSE_INPUT_DIR = r"D:\results\heywoodbess\psse\v0-0-4_sav_v0-0-4_dyr\DMAT"
        out_to_pkl(x86=True,PSSE_INPUT_DIR=PSSE_INPUT_DIR)

#     if not convert_psse_to_pkl:
#         PSCAD_INPUT_DIR = r"""C:\Grid\cg\results\pscad\vgrid"""
#         from cgbess.plotting.convert_to_pkl import psout_to_pkl
#         psout_to_pkl(x86=False,PSCAD_INPUT_DIR=PSCAD_INPUT_DIR)

