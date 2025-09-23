import sys
import os
import json

import pandas as pd 
import numpy as np
import math
import glob
import traceback
import ast
from pallet import load_specs_from_csv, load_specs_from_xlsx
from pallet import now_str
from typing import List, Optional, Callable

from pathlib import Path
from heywoodbess.appendices.appendix_generation_heywoodbess import generate_appendix_heywoodbess
from heywoodbess.appendices.replace_full_stops import replace_full_stops


APPENDIX_LIST = [
    {
        "plots-subdir": "Balanced Faults",
        "title": "DMAT 324 Benchmarking Balanced Faults",
    },
    {
        "plots-subdir": "TOV",
        "title": "DMAT 329 Benchmarking Temporary Overvoltage",
    },
    {
        "plots-subdir": "PFref Steps",
        "title": "DMAT 3210 Benchmarking Power Factor Reference Step Changes",
    },
    {
        "plots-subdir": "Qref Steps",
        "title": "DMAT 3210 Benchmarking Reactive Power Reference Step Changes",
    },
    {
        "plots-subdir": "Vref Steps",
        "title": "DMAT 3210 Benchmarking Voltage Reference Step Changes",
    },
    {
        "plots-subdir": "Pref Steps",
        "title": "DMAT 3211 Benchmarking Active Power References Changes",
    },
    {
        "plots-subdir": "Fgrid Steps",
        "title": "DMAT 3212 Benchmarking Grid Frequency Changes",
    },
    {
        "plots-subdir": "Vgrid Steps",
        "title": "DMAT 3214 Benchmarking Grid Voltage Changes",
    },
    {
        "plots-subdir": "Phase Step",
        "doc-number": "GRID-LINK-DOC-NUMBER-009",
        "title": "DMAT 3216 Benchmarking Phase Angle Changes",
    },
]


def create_appendix_benchmarking(
        PLOT_RESULTS_DIR,
        APPENDIX_OUTPUT_DIR,
        SEPARATE_CRG_DISCRG_APPENDIX:Optional[bool],
    ):
    print("Creating appendices ...")

    def caption_preprocessor(caption : str):
        return caption.replace("_", " ")
    
    doc_number = 0
    record_errors = []
    clauses = [f for f in os.listdir(PLOT_RESULTS_DIR)]

    for apdx in APPENDIX_LIST:
        if apdx["plots-subdir"] in clauses:
            print(apdx["plots-subdir"])
            doc_number = doc_number + 1
        
            replace_full_stops(os.path.join(PLOT_RESULTS_DIR,apdx["title"]))

            if SEPARATE_CRG_DISCRG_APPENDIX and not None:
                for IS_CHARGING_PRESENT in [True,False]:
                    if IS_CHARGING_PRESENT:
                        try:
                            plot_path = os.path.join(PLOT_RESULTS_DIR,apdx["plots-subdir"])
                            title_crg = f"{apdx['title']} - Charging"
                            output_path = os.path.join(APPENDIX_OUTPUT_DIR,"Charging Appendix",f"{title_crg}.pdf")

                            generate_appendix_heywoodbess(
                                project_name="heywoodbess",
                                client="Enzen",
                                title = title_crg,
                                doc_number=f"GRID-LINK-DOC-NUMBER-{doc_number}",
                                issued_date="22 April 2025",
                                revision="Rev 0",
                                revision_history_csv_path="revisionhistory.csv",
                                plots_directory=plot_path,
                                output_path=output_path,
                                bess_charging=IS_CHARGING_PRESENT,
                                caption_preprocessor_fn=caption_preprocessor,
                                )
                        except Exception as e:
                            traceback.print_exc()
                            print(f"Error: {e} --> CHARGING APPENDIX FAILED")
                            record_errors.append(f"{apdx['title']} - Charging Appendix -> Error: {e}")
                    else:
                        try:
                            plot_path = os.path.join(PLOT_RESULTS_DIR,apdx["plots-subdir"])
                            title_dcrg = apdx["title"]
                            output_path = os.path.join(APPENDIX_OUTPUT_DIR,"Discharging Appendix",f"{title_dcrg}.pdf")

                            generate_appendix_heywoodbess(
                                project_name="heywoodbess",
                                client="Enzen",
                                title = title_dcrg,
                                doc_number=f"GRID-LINK-DOC-NUMBER-{doc_number}",
                                issued_date="22 April 2025",
                                revision="Rev 0",
                                revision_history_csv_path="revisionhistory.csv",
                                plots_directory=plot_path,
                                output_path=output_path,
                                bess_charging=IS_CHARGING_PRESENT,
                                caption_preprocessor_fn=caption_preprocessor,
                                )

                        except Exception as e:
                            print(f"Error: {e} --> CHARGING APPENDIX FAILED")
                            record_errors.append(f"{apdx['title']} - Appendix -> Error: {e}")
            else:
                try:
                    plot_path = os.path.join(PLOT_RESULTS_DIR,apdx["plots-subdir"])
                    title = f"{apdx['title']}"
                    output_path = os.path.join(APPENDIX_OUTPUT_DIR,f"{title}.pdf")


                    generate_appendix_heywoodbess(
                        project_name="heywoodbess",
                        client="Enzen",
                        title = title,
                        doc_number=f"GRID-LINK-DOC-NUMBER-{doc_number}",
                        issued_date="22 April 2025",
                        revision="Rev 0",
                        revision_history_csv_path="revisionhistory.csv",
                        plots_directory=plot_path,
                        output_path=output_path,
                        bess_charging=None,
                        caption_preprocessor_fn=caption_preprocessor,
                        )
                except Exception as e:
                    print(f"Error: {e} --> COMBINED APPENDIX FAILED")
                    record_errors.append(f"{apdx['title']} - Charging Appendix -> Error: {e}")

    print("FINISHED APPENDIX GENERATION\n")
    print("All the Errors Encountered:")
    if len(record_errors)==0:
        print("No errors")
    else:
        for idx,error in enumerate(record_errors):
            print(f"{idx+1}. {error}\n")
    return

