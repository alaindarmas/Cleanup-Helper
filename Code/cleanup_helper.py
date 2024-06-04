import os
import shutil
import pandas as pd

# Load the configuration file
config_path = os.path.join(os.path.dirname(__file__), '../Documentation/ConfigFile_Helper.xlsx')
config_df = pd.read_excel(config_path)

# Convert the configuration DataFrame to a dictionary
config = {row['Key']: row['Value'] for _, row in config_df.iterrows()}

# Function to get the absolute path
def get_abs_path(relative_path):
    return os.path.abspath(os.path.join(os.path.dirname(__file__), relative_path))

# Define the paths
paths_to_clean = [
    get_abs_path(config['main_file_dir']),
    get_abs_path(config['filter_1_dir']),
    get_abs_path(config['filter_2_dir']),
    get_abs_path(config['filter_3_dir']),
    get_abs_path(config['text_files_dir'])
]

# Define the output directory path
output_dir = get_abs_path(config['output_excel_path'])

# Function to delete files and directories
def clean_directory(path):
    if os.path.exists(path):
        for root, dirs, files in os.walk(path):
            for file in files:
                os.remove(os.path.join(root, file))
            for dir in dirs:
                shutil.rmtree(os.path.join(root, dir))

# Delete existing files in specified directories
for path in paths_to_clean:
    clean_directory(path)

# Delete existing Excel files in the output directory
if os.path.exists(output_dir):
    for file in os.listdir(output_dir):
        if file.endswith('.xlsx'):
            os.remove(os.path.join(output_dir, file))

print("Cleanup complete.")
