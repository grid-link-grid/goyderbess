import pandas as pd

import re




def parse_logfile_dstates(log_file):
    # CHECK FOR DSTATE ERRORS
    with open(log_file, "r") as file:
        text = file.read()

    lines = text.splitlines()
    start_index = next(i for i, line in enumerate(lines) if "INITIAL CONDITIONS SUSPECT:" in line) + 1
    end_index = next(i for i, line in enumerate(lines) if "Channel output file is" in line) - 1

    # Extract column names and data
    columns = ["I", "DSTATE(I)","STATE(I)","MODEL","STATE","BUS-SCT","NAMEBASKV ID"] #re.split(r'\s{2,}', lines[start_index].strip())  # Split on 2+ spaces
    data_lines = lines[start_index + 1:end_index]  # Remaining lines after the header
    #print(data_lines)

    # Parse the data lines
    data = [re.split(r'\s{1,}', line.strip()) for line in data_lines if line.strip()]
    #print(data)
    # Process each child list
    for i, child in enumerate(data):
        if len(child) > 6:
            # Join the remaining items into a single 7th item
            new_item = ' '.join(map(str, child[6:]))  # Join items from index 6 onwards
            data[i] = child[:6] + [new_item]  # Replace child list with first 6 items + new 7th item

    # Create a DataFrame
    df = pd.DataFrame(data, columns=columns)

    # Convert numeric columns
    numeric_columns = ["I", "DSTATE(I)", "STATE(I)", "BUS-SCT"]
    for col in numeric_columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df