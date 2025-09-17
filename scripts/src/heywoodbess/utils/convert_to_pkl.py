import os
import sys

from typing import Optional, Union, Dict, List, Callable

# convert_psse_to_pkl = True #Set to true if you want to convert psse out files to pkls. 


# psout_to_pkl(x86=False,PSCAD_INPUT_DIR=PSCAD_INPUT_DIR)

def convert_to_pkl(convert_psse_to_pkl:bool,
                   PSSE_INPUT_DIR:Optional[str]=None,
                   PSCAD_INPUT_DIR:Optional[str]=None):


    if convert_psse_to_pkl:
        from heywoodbess.plotting.convert_to_pkl import out_to_pkl
        if PSSE_INPUT_DIR is not None:
            out_to_pkl(x86=True,PSSE_INPUT_DIR=PSSE_INPUT_DIR)
        else:
            print("Did not provide PSSE INPUT DIR for conversion!")

    if not convert_psse_to_pkl:
        from heywoodbess.plotting.convert_to_pkl import psout_to_pkl
        if PSCAD_INPUT_DIR is not None:

            psout_to_pkl(x86=False,PSCAD_INPUT_DIR=PSCAD_INPUT_DIR)
        else:
            print("Did not provide PSCAD INPUT DIR for pkl converesion!")
