# Author: jakub-kuba

import pandas as pd
import numpy as np
from datetime import datetime
import calendar
import sys
import os

def find_unique_file(subfolder_name, extension):
    """Finds file names in the folders"""
    subfolder_path = os.path.join(os.getcwd(), subfolder_name)

    # check if subfolder exists
    if not os.path.exists(subfolder_path) or not os.path.isdir(subfolder_path):
        print(f"Error: Subfolder '{subfolder_name}' does not exist.")
        sys.exit()

    # find all files with specific extension
    all_files = [x for x in os.listdir(subfolder_name) if
                  extension in x and 'crdownload' not in x and x[0] != '~'
                  and not x.startswith('.~lock')]
    
    # check the number of files with specific extension
    if len(all_files) == 0:
        print(f"Error: No files in '{subfolder_name}' subfolder!")
        sys.exit()
    if len(all_files) > 1:
        print(f"Error: There must be only one file in '{subfolder_name}' subfolder!")
        sys.exit()
    
    return subfolder_name+all_files[0]


def main():

    # list of all month abr.
    ALL_MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    # names of the files with source data
    SOURCE_FOLDER = "source file/"
    MAPPING_FOLDER = "mapping/"

    # file for missing programs
    NO_PROGRAMS = "missing programs.txt"

    source_file = find_unique_file(SOURCE_FOLDER, ".xlsx")
    mapping_file = find_unique_file(MAPPING_FOLDER, ".xlsx")
    
    # read all sheets from mapping tables file
    programs = pd.read_excel(mapping_file, sheet_name="Programs")
    cycles = pd.read_excel(mapping_file, sheet_name="Cycles")
    columns = pd.read_excel(mapping_file, sheet_name="Columns")

    #create necessary dictionaries and lists
    prog_dict = dict(zip(programs['Program'], programs['Full Program']))
    cycle_dict = dict(zip(cycles['Month'], cycles['Cycle']))

    source_columns = list(columns['Required Source Columns'].dropna())
    renamed_columns = list(columns['Renamed Columns'].dropna())
    data_types = list(columns['Data Type'].dropna())

    dtype_dict = dict(zip(source_columns, data_types))
    columns_changed = dict(zip(source_columns, renamed_columns))

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

    # rename source columns
    df = df.rename(columns=columns_changed)

    # round numbers to 6 decimal places
    df = df.round(6)

    # # remove "_" from column names
    # rename_dict = {col: col.rstrip("_") for col in df.columns if col.endswith("_")}
    # df = df.rename(columns=rename_dict)

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
    df = df.dropna(subset=ALL_MONTHS, how='all')

    # make a copy of df
    df_opposite = df.copy()

    # change all numbers in month columns to inverse number
    df_opposite[ALL_MONTHS] = -df_opposite[ALL_MONTHS]

    # replace Programs with blanks in the original df
    df['ProgramName'] = np.nan

    # concatenate both dataframes
    full = pd.concat([df_opposite, df])

    # check if there are ny missing progframs
    source_programs = full['ProgramName'].dropna().unique().tolist()
    missing_programs = list(source_programs - prog_dict.keys())

    # if there are missing programs, add their names in txt file
    if missing_programs:
        print(f'\nMissing programs found! Please check "{NO_PROGRAMS}" and update the mapping table.')
        with open(NO_PROGRAMS, 'w') as f:
            for line in missing_programs:
                f.write(f"{line}\n")
        sys.exit()
    else:
        print("\nNo missing programs")

    # replace programs where necessary (based on programs mapping table)
    full['ProgramName'] = full['ProgramName'].map(prog_dict)

    # create an empty list for DataFrames
    list_of_dfs = []

    # go through all month columns and get amount for each
    for x in ALL_MONTHS:
        new_df = full.copy()
        new_df['Amount'] = new_df[x]
        new_df['Year'] = new_df['Year'].dt.strftime('%Y')
        new_df['Month Name'] = x
        new_df = new_df[new_df[x].notna()]
        new_df = new_df.drop(ALL_MONTHS, axis=1)
        list_of_dfs.append(new_df)

    # concatenate all dataframes present in list_of_dfs
    final_df = pd.concat(list_of_dfs, ignore_index=True)

    # create a month column
    final_df['Month'] = "01 " + final_df['Month Name'] + " " + final_df['Year']

    # create a Month Number column based on first future chosen
    final_df['Month Number'] = first_future.month

    # create a cycle column based on month number and cycle dict
    final_df['Cycle'] = final_df['Month Number'].map(cycle_dict)

    # create a comment colummn
    final_df['Comment'] = "Reshuffle " + final_df['Cycle']

    # all all required columns
    final_df = final_df.reindex(final_df.columns.union(final_columns,sort=False),
                                axis=1, fill_value=np.nan)
    
    # limit DataFrame to columns required in final template and put them in correct folder
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