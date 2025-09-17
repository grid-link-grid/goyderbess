
import sys
import os
import json
import pandas as pd 
import numpy as np
import math
import glob




def caption_preprocessor(caption : str):
    return caption.replace("_", " ")