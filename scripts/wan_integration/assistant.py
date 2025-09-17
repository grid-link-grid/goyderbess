import sys

from pallet.utils.time_utils import now_str

# Import PF before any other internal modules so they don't need to change the path
PSSBIN_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSBIN"""
PSSPY_PATH = r"""C:\Program Files (x86)\PTI\PSSE34\PSSPY39"""
sys.path.append(PSSBIN_PATH)
sys.path.append(PSSPY_PATH)
import psse34
import psspy

import os
import pandas as pd

from gridlink.utils.wan.assistant.WanAssistant import WanAssistant


assistant = WanAssistant()
assistant.run()
