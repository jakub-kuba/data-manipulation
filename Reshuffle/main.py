# Author: jakub-kuba

import pandas as pd
import numpy as np
from datetime import datetime
import calendar
import sys
import os
import glob

def find_unique_file(subfolder_name, extension):
    """Finds file names in the folders"""
    subfolder_path = os.path.join(os.getcwd(), subfolder_name)

    # check if subfolder exists
    if not os.path.exists(subfolder_path) or not os.path.isdir(subfolder_path):
        print(f"Error: Subfolder '{subfolder_name}' does not exist.")
        sys.exit()
    # find all files with specific extension
    files = glob.glob(os.path.join(subfolder_path, extension))
    # check the number of files with specific extension
    if len(files) == 0:
        print(f"Error: No files in '{subfolder_name}' subfolder!")
        sys.exit()
    if len(files) > 1:
        print(f"Error: There must be only one file in '{subfolder_name}' subfolder!")
        sys.exit()
    
    return files[0]


def main():

    # list of all month abr.
    all_months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    # names of the files with source data
    source_folder = "source file"
    mapping_folder = "mapping"

    source_file = find_unique_file(source_folder, "*.xlsx")
    mapping_file = find_unique_file(mapping_folder, "*.xlsx")

    # read all sheets from mapping tables file
    programs = pd.read_excel(mapping_file, sheet_name="Programs")
    cycles = pd.read_excel(mapping_file, sheet_name="Cycles")
    columns = pd.read_excel(mapping_file, sheet_name="Columns")

    #create necessary dictionaries and lists
    prog_dict = dict(zip(programs['Program'], programs['Full Program']))

    cycle_dict = dict(zip(cycles['Month'], cycles['Cycle']))

    source_columns = [x for x in columns['Source Columns'].tolist() if str(x) != 'nan']
    data_types = [x for x in columns['Data Type'].tolist() if str(x) != 'nan']
    dtype_dict = dict(zip(source_columns, data_types))
    additional_columns = [x for x in columns['Additional Columns'].tolist() if str(x) != 'nan']
    final_columns = list(columns['Final Columns'].dropna())

    # read raw file, use required columns and their data types
    raw_file = pd.read_excel(source_file,
                             sheet_name="Data",
                             usecols=source_columns,
                             dtype=dtype_dict)
    
    # input first forecast month-year
    first_future = input("\nEnter the FIRST FUTURE MONTH in the format [mm-YYYY]: ")

    try:
        first_future = datetime.strptime(first_future, '%m-%Y')
    except ValueError:
        print("\nIncorrect date format")
        sys.exit()

    df = raw_file.copy()

    # remove "_" from column names
    rename_dict = {col: col.rstrip("_") for col in df.columns if col.endswith("_")}
    df = df.rename(columns=rename_dict)

    # change Year to datetime data type
    df['Year'] = pd.to_datetime(df['Year'], format='%Y')

    # limit source to NPC and Monetized
    df = df[df['Source'].isin(['Database', 'CSV'])]

    # delete all rows with years greater than the current one
    df = df[df['Year'] <= first_future]

    # delete all amounts from future months
    for x in range(first_future.month,13):
        df[calendar.month_abbr[x]] = np.where(df['Year'] == str(first_future.year),
                                              np.nan,
                                              df[calendar.month_abbr[x]])
    
    # drop rows where Result is empty
    df = df[df['Result'].notna()]

    # drop rows with blanks in all month columns
    df = df.dropna(subset=all_months, how='all')

    # make a copy of df
    df_opposite = df.copy()

    # change all numbers in month columns to inverse number
    df_opposite[all_months] = -df_opposite[all_months]

    # replace Programs with blanks in the original df
    df['Program'] = np.nan

    # concatenate both dataframes
    full = pd.concat([df_opposite, df])

    # replace programs where necessary (based on programs mapping table)
    full['Program'] = full['Program'].map(prog_dict)

    # fill Type column with blanks
    full['Type'] = np.nan

    # create an empty list for DataFrames
    list_of_dfs = []

    # go through all month columns and get amount for each
    for x in all_months:
        new_df = full.copy()
        new_df['Amount'] = new_df[x]
        new_df['My Year'] = new_df['Year'].dt.strftime('%Y')
        new_df['Month Name'] = x
        new_df = new_df[new_df[x].notna()]
        new_df = new_df.drop(all_months, axis=1)
        list_of_dfs.append(new_df)

    # concatenate all dataframes present in list_of_dfs
    final_df = pd.concat(list_of_dfs, ignore_index=True)

    # create a month column
    final_df['Month'] = "01 " + final_df['Month Name'] + " " + final_df['My Year']

    # create a Month Number column based on first future chosen
    final_df['Month Number'] = first_future.month

    # create a cycle column based on month number and cycle dict
    final_df['Cycle'] = final_df['Month Number'].map(cycle_dict)

    # create a comment colummn
    final_df['Comment'] = "Reshuffle " + final_df['Cycle']

    # rename a few columns to match the final template
    final_df = final_df.rename(columns={"Code 1": "Code1",
                                        "Code 2": "Code2",
                                        "Code 3": "Code3",
                                        "Program": "ProgramName"}
                                        )
    
    # # add empty columns required for final template
    final_df = pd.concat([final_df, pd.DataFrame(columns=additional_columns)])

    # limit dataframe to columns required in the final template and put them in correct order
    final_df = final_df[final_columns]

    # create a dataframe for additional worksheet
    next = pd.DataFrame(columns=['Id'])

    now = datetime.now()
    date_now = now.strftime('%d-%b-%Y_%H-%M-%S')

    with pd.ExcelWriter('results/reshuffle_'+date_now+'.xlsx') as writer:
        final_df.to_excel(writer, sheet_name='Upload', index=False)
        next.to_excel(writer, sheet_name='Next', index=False)

    print("\nJob Complete!")


if __name__ == "__main__":
    main()