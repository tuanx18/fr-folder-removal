# Original Developer: TuanHA47

import pandas as pd
import os
import tkinter as tk
from tkinter import Scrollbar
from tkinter import messagebox
import openpyxl
import time
import shutil

# ------------------------------------------------------------ #
# Input the path for Excel template here
excel_path = "input/DPR-test.xlsx"
origin_slk_path = "fr-det-dlk2-safelake-pipelines-testing"
origin_dlk_path = "fr-det-dlk2-pipelines-testing"
# ------------------------------------------------------------ #

# Start timer
start_timer = time.time()

# Get output name
parts = excel_path.split("/")
excel_original_name = parts[-1]
raw_name = parts[-1].replace(".xlsx", "")

# Read the Excel
df = pd.read_excel(excel_path, header=None)

# Find the row indices containing "//"
indices_to_remove = df.index[df.apply(lambda row: '//' in row.astype(str), axis=1)]

# If any row contains "//", remove it and all rows below it
if not indices_to_remove.empty:
    last_index_to_remove = indices_to_remove.max()
    df = df.iloc[:last_index_to_remove]

# Extract data from specific range after removing rows
data = df.iloc[2:100, 0:9].values  # Adjust the range as needed
lendf = len(data)

# Convert list elements to strings and replace existing commas with "+"
data = [[str(item).replace(',', '+') if isinstance(item, str) else item for item in row] for row in data]

# Join each row with commas
data = [','.join(map(str, row)) for row in data]

# Find the index of the first item that contains "//"
index_of_double_slash = next((index for index, item in enumerate(data) if "//," in item), None)

# Remove all items after the first "//" item, including itself
if index_of_double_slash is not None:
    data = data[:index_of_double_slash]

process_data_list = []

# To delete
all_delete_file = []                # Store the file path
all_delete_folder_name = []         # Store the folder name
all_delete_folder_path = []         # Store the folder path

# Master Dictionary: To store later Excel export data
master_dictionary = {}
master_tags = []
master_lakes = []

# Function to check if a file name contains any tag
def contains_tag(file_name):
    return all(tag in file_name.lower() for tag in tags)

# Function to check if a folder name contains all tags
def contains_all_tags(folder_name):
    return all(tag in folder_name.lower() for tag in tags)

# Function to get the full path of a directory
def get_full_path(directory_name):
    return os.path.join(new_path, directory_name)

def get_full_path_acltd(directory_name):
    return os.path.join(acltd_path, directory_name)

def get_full_path_clsd(directory_name):
    return os.path.join(clsd_path, directory_name)

def get_full_path_dnorm(directory_name):
    return os.path.join(dnorm_path, directory_name)

# Function to get all items in a directory
def get_directory_items(directory_path):
    return [os.path.join(directory_path, item) for item in os.listdir(directory_path)]

# Function: change all \, \\ to / for all paths from a list
def normalize_path(lst):
    result = []
    for item in lst:
        modded_item = item.replace('\\\\', '\\')
        modded_item_2 = modded_item.replace('\\', '/')
        modded_item_3 = modded_item_2.replace('//', '/')
        result.append(modded_item_3)
    return result

# Function: remove duplicates from a list
def remove_duplicates(lst):
    result = []
    for item in lst:
        if item not in result:
            result.append(item)
    return result

def process_list(lst1, lst2, lst3):
    lst1 = normalize_path(lst1)
    lst2 = normalize_path(lst2)
    lst3 = normalize_path(lst3)

    lst1 = remove_duplicates(lst1)
    lst2 = remove_duplicates(lst2)
    lst3 = remove_duplicates(lst3)

    return lst1, lst2, lst3


# ### Alter this with desire list to be process
# process_data_list.append(data[15])                   # Test, CUT list of data
################################ PROCEED
process_data_list.extend(data)
################################ PROCEED

# Proceed for the whole list
item_index = 0
for record in process_data_list:

    # To delete dictionary for each item
    dr_dictionary = {"folders": [], "files": []}
    edp_dictionary = {"folders": [], "files": []}
    r_dictionary = {"folders": [], "files": []}

    # Test list
    test_row = record
    data_list = test_row.split(',')

    # Split region => turn into a list
    data_list[5] = data_list[5].replace(" ", "").split('+')

    # Change all value from col 7-8-9 to 0 if not a correct value
    if str(data_list[6]) != "0" and str(data_list[6]) != "1":
        data_list[6] = 0

    if str(data_list[7]) != "0" and str(data_list[7]) != "1":
        data_list[7] = 0

    if str(data_list[8]) != "0" and str(data_list[8]) != "1":
        data_list[8] = 0

    data_list[6] = int(data_list[6])
    data_list[7] = int(data_list[7])
    data_list[8] = int(data_list[8])

    # Namespace
    namespace = data_list[3].lower()

    tags = [data_list[0].lower(), data_list[1].lower(), data_list[2].lower()]
    tags = [item for item in tags if item != "nan"]
    master_tags.append(tags)

    # Get Key name for the master dictionary item
    item_index += 1
    key_name = f"item_{item_index}"
    master_dictionary[key_name] = data_list

    # Master SLK/DLK
    if data_list[4].lower() == "slk":
        master_lakes.append("SLK")

    elif data_list[4].lower() == "dlk":
        master_lakes.append("DLK")

    else:
        master_lakes.append("Undefined Lake")

    # Case 1: SLK - dataset-registry
    if data_list[4].lower() == "slk" and int(data_list[6]) == 1:
        path = f"{origin_slk_path}/dataset-registry"     # [1] Change when finish testing - SLK path

        for region in data_list[5]:

            # Reset Temp lists
            to_delete_file_list = []
            to_delete_folder_list = []
            to_delete_folder_path_list = []

            region_lower = region.lower()
            new_path = f"{path}/{region_lower}/{namespace}"

            # Traverse the directory structure starting from new_path
            for root, dirs, files in os.walk(new_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    if contains_tag(file):
                        to_delete_file_list.append(file_path)

            # Add to master lists + dictionary
            all_delete_file.extend(to_delete_file_list)
            all_delete_folder_name.extend(to_delete_folder_list)
            all_delete_folder_path.extend(to_delete_folder_path_list)
            all_delete_file, all_delete_folder_name, all_delete_folder_path = process_list(all_delete_file, all_delete_folder_name, all_delete_folder_path)

            dr_dictionary['folders'].extend(to_delete_folder_path_list)
            dr_dictionary['files'].extend(to_delete_file_list)

    # Case 2: SLK - event-driven-pipelines
    if data_list[4].lower() == "slk" and int(data_list[7]) == 1:
        path = f"{origin_slk_path}/event-driven-pipelines"

        # Split Regions if there is more than 1
        for region in data_list[5]:

            # Reset Temp lists
            to_delete_file_list = []
            to_delete_folder_list = []
            to_delete_folder_path_list = []

            region_lower = region.lower()
            new_path = f"{path}/{region_lower}/{namespace}"

            # Get folder name using matched TAGs
            for root, dirs, files in os.walk(new_path):
                for directory in dirs:
                    if contains_all_tags(directory):
                        to_delete_folder_list.append(directory)

            # Get FULL PATH of folder
            for root, dirs, files in os.walk(new_path):
                for directory in dirs:
                    if contains_all_tags(directory):
                        directory_path = get_full_path(directory)
                        to_delete_folder_path_list.append(directory_path)

            # Traverse the directory structure starting from new_path
            for directory_name in to_delete_folder_list:
                directory_path = get_full_path(directory_name)
                directory_items = get_directory_items(directory_path)
                to_delete_file_list.extend(directory_items)

            # Add to master lists
            all_delete_file.extend(to_delete_file_list)
            all_delete_folder_name.extend(to_delete_folder_list)
            all_delete_folder_path.extend(to_delete_folder_path_list)
            all_delete_file, all_delete_folder_name, all_delete_folder_path = process_list(all_delete_file,
                                                                                           all_delete_folder_name,
                                                                                           all_delete_folder_path)

            edp_dictionary['folders'].extend(to_delete_folder_path_list)
            edp_dictionary['files'].extend(to_delete_file_list)


    # Case 3: SLK - resources
    if data_list[4].lower() == "slk" and int(data_list[8]) == 1:
        path = f"{origin_slk_path}/resources"

        for region in data_list[5]:

            # Reset Temp lists
            to_delete_file_list = []
            to_delete_folder_list = []
            to_delete_folder_path_list = []

            region_lower = region.lower()
            new_path = f"{path}/{region_lower}/bigquery/{namespace}/module/DDL"

            # Split the new_path to 3 layer folders: "accumulated" and "cleansed"
            acltd_path = f"{new_path}/accumulated"
            clsd_path = f"{new_path}/cleansed"
            dnorm_path = f"{new_path}/denormalized"

            # Split case 1 : accumulated - find all folders that match tags
            if os.path.exists(acltd_path):

                for root, dirs, files in os.walk(acltd_path):
                    for directory in dirs:
                        if contains_all_tags(directory):
                            to_delete_folder_list.append(directory)

                # Get FULL PATH of folder
                for root, dirs, files in os.walk(acltd_path):
                    for directory in dirs:
                        if contains_all_tags(directory):
                            directory_path = get_full_path_acltd(directory)
                            to_delete_folder_path_list.append(directory_path)

                # Traverse the directory structure starting from new_path
                for directory_name in to_delete_folder_list:
                    directory_path = get_full_path_acltd(directory_name)
                    directory_items = get_directory_items(directory_path)
                    to_delete_file_list.extend(directory_items)

                # Add to master lists
                all_delete_file.extend(to_delete_file_list)
                all_delete_folder_name.extend(to_delete_folder_list)
                all_delete_folder_path.extend(to_delete_folder_path_list)
                all_delete_file, all_delete_folder_name, all_delete_folder_path = process_list(all_delete_file,
                                                                                               all_delete_folder_name,
                                                                                               all_delete_folder_path)

                r_dictionary['folders'].extend(to_delete_folder_path_list)
                r_dictionary['files'].extend(to_delete_file_list)

            else:
                pass

            # Split case 2 : cleansed - find all folders that match tags
            # Reset Temp lists
            to_delete_file_list = []
            to_delete_folder_list = []
            to_delete_folder_path_list = []

            # Only open in case there is such a path exist
            if os.path.exists(clsd_path):

                for root, dirs, files in os.walk(clsd_path):
                    for directory in dirs:
                        if contains_all_tags(directory):
                            to_delete_folder_list.append(directory)

                # Get FULL PATH of folder
                for root, dirs, files in os.walk(clsd_path):
                    for directory in dirs:
                        if contains_all_tags(directory):
                            directory_path = get_full_path_clsd(directory)
                            to_delete_folder_path_list.append(directory_path)

                # Traverse the directory structure starting from new_path
                for directory_name in to_delete_folder_list:
                    directory_path = get_full_path_clsd(directory_name)
                    directory_items = get_directory_items(directory_path)
                    to_delete_file_list.extend(directory_items)

                # Add to master lists
                all_delete_file.extend(to_delete_file_list)
                all_delete_folder_name.extend(to_delete_folder_list)
                all_delete_folder_path.extend(to_delete_folder_path_list)
                all_delete_file, all_delete_folder_name, all_delete_folder_path = process_list(all_delete_file,
                                                                                               all_delete_folder_name,
                                                                                               all_delete_folder_path)

                r_dictionary['folders'].extend(to_delete_folder_path_list)
                r_dictionary['files'].extend(to_delete_file_list)

            else:
                pass

            # Split case 3 : denormalized - find all folders that match tags
            if os.path.exists(dnorm_path):

                # Reset Temp lists
                to_delete_file_list = []
                to_delete_folder_list = []
                to_delete_folder_path_list = []
                for root, dirs, files in os.walk(dnorm_path):
                    for directory in dirs:
                        if contains_all_tags(directory):
                            to_delete_folder_list.append(directory)

                # Get FULL PATH of folder
                for root, dirs, files in os.walk(dnorm_path):
                    for directory in dirs:
                        if contains_all_tags(directory):
                            directory_path = get_full_path_dnorm(directory)
                            to_delete_folder_path_list.append(directory_path)

                # Traverse the directory structure starting from new_path
                for directory_name in to_delete_folder_list:
                    directory_path = get_full_path_dnorm(directory_name)
                    directory_items = get_directory_items(directory_path)
                    to_delete_file_list.extend(directory_items)

                # Add to master lists
                all_delete_file.extend(to_delete_file_list)
                all_delete_folder_name.extend(to_delete_folder_list)
                all_delete_folder_path.extend(to_delete_folder_path_list)
                all_delete_file, all_delete_folder_name, all_delete_folder_path = process_list(all_delete_file,
                                                                                               all_delete_folder_name,
                                                                                               all_delete_folder_path)

                r_dictionary['folders'].extend(to_delete_folder_path_list)
                r_dictionary['files'].extend(to_delete_file_list)

            else:
                pass

    # Case 4: DLK - dataset-registry
    if data_list[4].lower() == "dlk" and int(data_list[6]) == 1:

        # Reset Temp lists
        to_delete_file_list = []
        to_delete_folder_list = []
        to_delete_folder_path_list = []

        path = f"{origin_dlk_path}/dataset-registry"
        new_path = f"{path}/{namespace}"

        # Traverse the directory structure starting from new_path
        for root, dirs, files in os.walk(new_path):
            for file in files:
                file_path = os.path.join(root, file)
                if contains_tag(file):
                    to_delete_file_list.append(file_path)

        # Case: confirm if delete all file on a folder -> delete folder
        split_path = to_delete_file_list.copy()
        counter_path = to_delete_file_list.copy()

        split_path = [s.replace("\\", "/") for s in split_path]     # Replace "\" with "/" for consistency
        split_path = ["/".join(path.split("/")[:-1]) for path in split_path]    # Get the paths without file names
        split_path = list(set(split_path))  # Remove duplicates

        path_counts = []    # Count path

        for path in split_path:
            if os.path.isdir(path):
                files = os.listdir(path)    # Get all items to store in a variable
                total_files = len(files)    # Count items
                path_counts.append([path, total_files])
            else:
                # If the path is not a directory, append None as the count
                path_counts.append([path, None])
                print(f"{path} directory is not found or not accessible.")

        _ = 0
        for f in split_path:

            count = 0
            for i in counter_path:
                i = i.replace("\\", "/")
                if f in i:
                    count += 1
            path_counts[_].append(count)
            _ += 1

        for case_set in path_counts:
            if case_set[1] == case_set[2]:
                to_delete_folder_path_list.append(case_set[0])

        # Add to master lists
        all_delete_file.extend(to_delete_file_list)
        all_delete_folder_name.extend(to_delete_folder_list)
        all_delete_folder_path.extend(to_delete_folder_path_list)
        all_delete_file, all_delete_folder_name, all_delete_folder_path = process_list(all_delete_file,
                                                                                       all_delete_folder_name,
                                                                                       all_delete_folder_path)

        dr_dictionary['folders'].extend(to_delete_folder_path_list)
        dr_dictionary['files'].extend(to_delete_file_list)

    # Case 5: DLK - event-driven-pipelines
    if data_list[4].lower() == "dlk" and int(data_list[7]) == 1:
        # Reset Temp lists
        to_delete_file_list = []
        to_delete_folder_list = []
        to_delete_folder_path_list = []

        path = f"{origin_dlk_path}/event-driven-pipelines"
        new_path = f"{path}/{namespace}"

        # Get folder name using matched TAGs
        for root, dirs, files in os.walk(new_path):
            for directory in dirs:
                if contains_all_tags(directory):
                    to_delete_folder_list.append(directory)

        # Get FULL PATH of folder
        for root, dirs, files in os.walk(new_path):
            for directory in dirs:
                if contains_all_tags(directory):
                    directory_path = get_full_path(directory)
                    to_delete_folder_path_list.append(directory_path)

        # Traverse the directory structure starting from new_path
        for directory_name in to_delete_folder_list:
            directory_path = get_full_path(directory_name)
            directory_items = get_directory_items(directory_path)
            to_delete_file_list.extend(directory_items)

        # Add to master lists
        all_delete_file.extend(to_delete_file_list)
        all_delete_folder_name.extend(to_delete_folder_list)
        all_delete_folder_path.extend(to_delete_folder_path_list)
        all_delete_file, all_delete_folder_name, all_delete_folder_path = process_list(all_delete_file,
                                                                                       all_delete_folder_name,
                                                                                       all_delete_folder_path)

        edp_dictionary['folders'].extend(to_delete_folder_path_list)
        edp_dictionary['files'].extend(to_delete_file_list)

    # Case 6: DLK - resources
    if data_list[4].lower() == "dlk" and int(data_list[8]) == 1:

        # Reset Temp lists
        to_delete_file_list = []
        to_delete_folder_list = []
        to_delete_folder_path_list = []

        path = f"{origin_dlk_path}/resources"
        new_path = f"{path}/bigquery/{namespace}/module/DDL"

        # Split the new_path to 3 layer folders: "accumulated" and "cleansed"
        acltd_path = f"{new_path}/accumulated"
        clsd_path = f"{new_path}/cleansed"
        dnorm_path = f"{new_path}/denormalized"

        # Split case 1 : accumulated - find all folders that match tags
        if os.path.exists(acltd_path):

            for root, dirs, files in os.walk(acltd_path):
                for directory in dirs:
                    if contains_all_tags(directory):
                        to_delete_folder_list.append(directory)

            # Get FULL PATH of folder
            for root, dirs, files in os.walk(acltd_path):
                for directory in dirs:
                    if contains_all_tags(directory):
                        directory_path = get_full_path_acltd(directory)
                        to_delete_folder_path_list.append(directory_path)

            # Traverse the directory structure starting from new_path
            for directory_name in to_delete_folder_list:
                directory_path = get_full_path_acltd(directory_name)
                directory_items = get_directory_items(directory_path)
                to_delete_file_list.extend(directory_items)

            # Add to master lists
            all_delete_file.extend(to_delete_file_list)
            all_delete_folder_name.extend(to_delete_folder_list)
            all_delete_folder_path.extend(to_delete_folder_path_list)
            all_delete_file, all_delete_folder_name, all_delete_folder_path = process_list(all_delete_file,
                                                                                           all_delete_folder_name,
                                                                                           all_delete_folder_path)

            r_dictionary['folders'].extend(to_delete_folder_path_list)
            r_dictionary['files'].extend(to_delete_file_list)

        else:
            pass

        # Split case 2 : cleansed - find all folders that match tags
        # Reset Temp lists
        to_delete_file_list = []
        to_delete_folder_list = []
        to_delete_folder_path_list = []

        # Only open in case there is such a path exist
        if os.path.exists(clsd_path):

            for root, dirs, files in os.walk(clsd_path):
                for directory in dirs:
                    if contains_all_tags(directory):
                        to_delete_folder_list.append(directory)

            # Get FULL PATH of folder
            for root, dirs, files in os.walk(clsd_path):
                for directory in dirs:
                    if contains_all_tags(directory):
                        directory_path = get_full_path_clsd(directory)
                        to_delete_folder_path_list.append(directory_path)

            # Traverse the directory structure starting from new_path
            for directory_name in to_delete_folder_list:
                directory_path = get_full_path_clsd(directory_name)
                directory_items = get_directory_items(directory_path)
                to_delete_file_list.extend(directory_items)

            # Add to master lists
            all_delete_file.extend(to_delete_file_list)
            all_delete_folder_name.extend(to_delete_folder_list)
            all_delete_folder_path.extend(to_delete_folder_path_list)
            all_delete_file, all_delete_folder_name, all_delete_folder_path = process_list(all_delete_file,
                                                                                           all_delete_folder_name,
                                                                                           all_delete_folder_path)

            r_dictionary['folders'].extend(to_delete_folder_path_list)
            r_dictionary['files'].extend(to_delete_file_list)

        else:
            pass

        # Split case 3 : denormalized - find all folders that match tags
        if os.path.exists(dnorm_path):

            # Reset Temp lists
            to_delete_file_list = []
            to_delete_folder_list = []
            to_delete_folder_path_list = []
            for root, dirs, files in os.walk(dnorm_path):
                for directory in dirs:
                    if contains_all_tags(directory):
                        to_delete_folder_list.append(directory)

            # Get FULL PATH of folder
            for root, dirs, files in os.walk(dnorm_path):
                for directory in dirs:
                    if contains_all_tags(directory):
                        directory_path = get_full_path_dnorm(directory)
                        to_delete_folder_path_list.append(directory_path)

            # Traverse the directory structure starting from new_path
            for directory_name in to_delete_folder_list:
                directory_path = get_full_path_dnorm(directory_name)
                directory_items = get_directory_items(directory_path)
                to_delete_file_list.extend(directory_items)

            # Add to master lists
            all_delete_file.extend(to_delete_file_list)
            all_delete_folder_name.extend(to_delete_folder_list)
            all_delete_folder_path.extend(to_delete_folder_path_list)
            all_delete_file, all_delete_folder_name, all_delete_folder_path = process_list(all_delete_file,
                                                                                           all_delete_folder_name,
                                                                                           all_delete_folder_path)

            r_dictionary['folders'].extend(to_delete_folder_path_list)
            r_dictionary['files'].extend(to_delete_file_list)

        else:
            pass

    if data_list[4].lower() != "dlk" and data_list[4].lower() != "slk":
        pass

    # Add to master dictionary
    master_dictionary[key_name].append(dr_dictionary)
    master_dictionary[key_name].append(edp_dictionary)
    master_dictionary[key_name].append(r_dictionary)

##### END OF LOOP

# Step: remove duplicates from to_delete_folder_path_list, to_delete_folder_list, to_delete_file_list
all_delete_folder_path = remove_duplicates(all_delete_folder_path)
all_delete_folder_name = remove_duplicates(all_delete_folder_name)
all_delete_file = remove_duplicates(all_delete_file)

# Step: normalize slashes
all_delete_folder_path = normalize_path(all_delete_folder_path)
all_delete_file = normalize_path(all_delete_file)

# Print the results
print("-" * 120)
print(f"Input records scanned: {len(process_data_list)}")
print(f"Unique Folders Found: {len(all_delete_folder_path)}")
print(f"Unique Files Found: {len(all_delete_file)}")

master_list = []
master_header = ["index", "item_id", "location", "file/folder", "path"]
master_list.append(master_header)
_ = 0

for key, value in master_dictionary.items():

    dr_folder = value[-3]["folders"]
    dr_file = value[-3]["files"]

    edp_folder = value[-2]["folders"]
    edp_file = value[-2]["files"]

    r_folder = value[-1]["folders"]
    r_file = value[-1]["files"]

    dr_folder, edp_folder, r_folder = process_list(dr_folder, edp_folder, r_folder)
    dr_file, edp_file, r_file = process_list(dr_file, edp_file, r_file)

    # Append to master list
    for item in dr_folder:
        _ += 1
        master_list.append([_, key, "dataset-registry", "folder", item])

    for item in dr_file:
        _ += 1
        master_list.append([_, key, "dataset-registry", "file", item])

    for item in edp_folder:
        _ += 1
        master_list.append([_, key, "event-driven-pipelines", "folder", item])

    for item in edp_file:
        _ += 1
        master_list.append([_, key, "event-driven-pipelines", "file", item])

    for item in r_folder:
        _ += 1
        master_list.append([_, key, "resources", "folder", item])

    for item in r_file:
        _ += 1
        master_list.append([_, key, "resources", "file", item])

### EXCEL PROCESSING
print("-" * 120)
print("Creating Excel file...")
# Output directory: Excel
output_dir = "output"
output_path = os.path.join(output_dir, f"{excel_original_name}")

# Ensure the output directory exists
os.makedirs(output_dir, exist_ok=True)

# Create a new Workbook
wb = openpyxl.Workbook()

# Select the active worksheet
ws = wb.active

# Write data to the worksheet
for row_idx, row_data in enumerate(master_list, 1):  # Start from row 1
    for col_idx, cell_value in enumerate(row_data, 1):  # Start from column 1
        ws.cell(row=row_idx, column=col_idx, value=cell_value)

# Save the workbook
wb.save(output_path)

if os.path.exists(output_path):
    print("Excel file saved successfully at:", output_path)
else:
    print("Failed creating Excel file.")
print("-" * 120)


### CSV PROCESSING
print("Creating Csv/Txt file...")

# Output directory: Text / CSV
r_output_dir = "output"
r_output_path = os.path.join(r_output_dir, f"{raw_name}.txt")

# Combine both lists
combined_list = all_delete_folder_path + all_delete_file

# Write data to the text file
with open(r_output_path, "w") as file:
    o = 1
    for item in combined_list:
        file.write(str(o) + "," + str(item) + "\n")
        o += 1

if os.path.exists(output_path):
    print("TXT/CSV file saved successfully at:", r_output_path)
else:
    print("Failed creating txt/csv file.")
print("-" * 120)

# --------------------------- Tkinter Interface -------------------------------- #
# Function to display list information
def display_list_info(lst):

    # Create a label to display the table
    item_id = 0
    current_item = " "
    for item in lst[1:]:
        if current_item != item[1]:
            current_item = item[1]
            if item_id == 0:
                item_text = f"{current_item}  -  {master_lakes[item_id]}  -  {master_tags[item_id]}"
                item_id += 1
            else:
                item_text = f"\n{current_item}  -  {master_lakes[item_id]}  -  {master_tags[item_id]}"
                item_id += 1

            item_label = tk.Label(scrollable_frame, text=item_text, wraplength=1200, justify="left", anchor="w",
                             font=("Times New Roman", 14, "bold"), fg="blue")
            item_label.pack(fill="both")

        text = f"{item[0]} ) {item[2]} > {item[3]} \n\t{item[4]}"

        label = tk.Label(scrollable_frame, text=text, wraplength=1200, justify="left", anchor="w",
                         font=("Times New Roman", 11))
        label.pack(fill="both")


# Button 2: Open files
def open_excel_file():
    if os.path.exists(output_path):
        os.startfile(output_path)
    else:
        print("Excel file does not exist.")


def open_txt_file():
    if os.path.exists(r_output_path):
        os.startfile(r_output_path)
    else:
        print("Text file does not exist.")


def delete_files():
    # Delete each file
    for file_path in combined_list:
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted file: {file_path}")
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
                print(f"Deleted folder: {file_path}")
            else:
                print(f"Invalid path: {file_path}")
        except Exception as e:
            print(f"\nError deleting {file_path}: {e}")


# Function to create a warning dialog
def show_warning_dialog():
    # Ask for confirmation
    confirm = messagebox.askyesno("Confirmation", f"Are you sure to delete the {len(combined_list)} files?")

    if confirm:
        delete_files()


# Function to calculate the center coordinates of the screen
def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_coordinate = (screen_width - width) // 2
    y_coordinate = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x_coordinate}+{y_coordinate}")

# Create Tkinter window
root = tk.Tk()
root.title("Verify File Removal")
window_width = 1444
window_height = 850
center_window(root, window_width, window_height)

# Add title label
title_label = tk.Label(root, text=f"{len(all_delete_folder_path) + len(all_delete_file)} Unique Items found ({len(all_delete_folder_path)} Folders and {len(all_delete_file)} Files/Sub-Folders)", font=("Times New Roman", 20))
title_label.pack(side=tk.TOP, pady=10)

# Create a frame for the scrolling area
scroll_frame = tk.Frame(root, width=1300, height=700)
scroll_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

# Create a canvas and scrollbar
canvas = tk.Canvas(scroll_frame, width=1300, height=696)
scrollbar = Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

# Configure scrollbar and canvas
scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

# Display information for each list
display_list_info(master_list)

# Pack canvas and scrollbar
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Create buttons
button1 = tk.Button(root, text="Refresh Contents", font=("Times New Roman", 14))
button2 = tk.Button(root, text="Open in Excel", font=("Times New Roman", 14))
button2b = tk.Button(root, text="Open in CSV", font=("Times New Roman", 14))
button3 = tk.Button(root, text="Delete All Selected", font=("Times New Roman", 14))

# Position buttons
button1.place(relx=0.1, rely=0.95, anchor=tk.CENTER)
button2.place(relx=0.45, rely=0.95, anchor=tk.CENTER)
button2b.place(relx=0.55, rely=0.95, anchor=tk.CENTER)
button3.place(relx=0.9, rely=0.95, anchor=tk.CENTER)

# Attach the function to the button
button2.config(command=open_excel_file)
button2b.config(command=open_txt_file)
button3.config(command=show_warning_dialog)

end_timer = time.time()
time_consumed = round(end_timer - start_timer, 3)
print(f"Time taken to process the program: {time_consumed} seconds")
print("-" * 120)

# Start the GUI event loop
root.mainloop()