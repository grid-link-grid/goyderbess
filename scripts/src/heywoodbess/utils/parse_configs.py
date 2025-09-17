import pandas as pd
import re
import os


import os
import re
import pandas as pd



def cfg_to_csv(cfg_path: str, results_dir: str):
    # Read the file content
    with open(cfg_path, 'r', encoding="utf-8") as file:
        lines = file.readlines()

    parsed_data = []
    section = None  # Track the current section
    
    pattern = r"^(\S+)\s*=\s*([-+]?[0-9]*\.?[0-9]+|\S+)\s*(?:#\s*(.*))?$"
    
    for line in lines:
        line = line.strip()
        
        if not line or line.startswith("#"):  # Skip empty lines and comments
            continue
        
        if line.startswith("[") and line.endswith("]"):  # Detect section headers
            section = line.strip("[]")
            continue
        
        match = re.match(pattern, line)
        if match:
            name, value, comment = match.groups()
            parsed_data.append({
                "section": section if section else "N/A",
                "name": name.strip(),
                "section.name": section+"."+name.strip() if section else "N/A",
                "value": value.strip(),
                "description": comment.strip() if comment else "N/A"
            })
        else:
            print(f"Line not matched: {line}")

    # Convert parsed data to a DataFrame and save as CSV
    df = pd.DataFrame(parsed_data, columns=["section", "name", "section.name","value", "description"])
    csv_name = os.path.basename(cfg_path).replace(".txt", ".csv")
    df.to_csv(os.path.join(results_dir, csv_name), index=False)


def r0_cfg_to_csv(cfg_path: str, results_dir: str):
    # Read the file content
    with open(cfg_path, 'r', encoding="utf-8") as file:
        lines = file.readlines()

    # Initialize a list to store parsed data
    parsed_data = []

    # Updated regex pattern
    pattern = r"^(\S+)\s+(-?\d*\.?\d*|[A-Za-z_]+)\s*(?:#\s*(.*))?$"

    # Loop through each line to parse
    for line in lines:
        line = line.strip()  # Remove leading and trailing whitespace
        if line.startswith("#") or not line:  # Skip comments or empty lines
            continue
        match = re.match(pattern, line)
        if match:
            name, value, description = match.groups()
            parsed_data.append({
                "name": name, 
                "value": value if value else "N/A", 
                "description": description.strip() if description else "N/A"
            })
        else:
            print(f"Line not matched: {line}")

    # Create the DataFrame
    csv_name = os.path.basename(cfg_path).replace(".txt",".csv") 
    df = pd.DataFrame(parsed_data, columns=["name", "value", "description"])
    df.to_csv(os.path.join(results_dir,os.path.join(csv_name)), index=False)


# Example usage
# cfg_to_csv("path/to/config.txt", "path/to/results_dir")



def dyr_to_csv(dyr_path: str, results_dir: str):
    # Initialize a list to store the name-value pairs
    parsed_data = []

    # Read the file content
    with open(dyr_path, 'r') as file:
        lines = file.readlines()

    # Processing the file
    setting_names = []  # To temporarily store setting names
    for line in lines:
        line = line.strip()
        if line.startswith("@!/"):  # Identify setting name rows
            setting_names = re.split(r'\s{2,}|\t', line[3:].strip())  # Split on 2+ spaces or tabs
        elif setting_names:  # Process values if setting names were defined
            values = re.split(r'\s{2,}|\t', line.strip())  # Split on 2+ spaces or tabs
            if len(values) == len(setting_names):  # Ensure matching length
                for name, value in zip(setting_names, values):
                    parsed_data.append({"name": name, "value": value})

    # Convert the list of dictionaries to a DataFrame
    df = pd.DataFrame(parsed_data, columns=["name", "value"])
    output_file = os.path.basename(dyr_path).replace(".dyr",".csv")
    # Save to CSV
    df.to_csv(os.path.join(results_dir,output_file), index=False)

