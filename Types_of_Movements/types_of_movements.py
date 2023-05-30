# Author: jakub-kuba

import numpy as np
import pandas as pd
import datetime
import dateutil.relativedelta
import sys
import os


#columns required
columns = ['ID',
        'GON',
        'Continent',
        'Country',
        'City',
        'Branch',
        'Specialization',
        'Dev Center',
        'State',
        'PDG',
        'Payment Center',
        'Office',
        'Resource Type',
        'Type of WTE',
        'Level',
        'Company',
        'WTE',
        'Date',
        'Program',
        'Agreement']

dtype_dict = {'ID' : str,
            'GON' : str,
            'Dev Center' : str,
            'State' : str,
            'PDG' : str,
            'Payment Center' : str,
            'Office' : str,
            'Resource Type' : str,
            'Type of WTE' : str,
            'Level' : str,
            'Company' : str,
            'Program' : str,
            'Agreement' : str,
            }

#required listsL=:
exterior_leased = ['Outer Personnel (Leased - Fixed Price)',
                'Outer Personnel (Leased - Time & Materials)',
                'Outer Personnel (Leased - Consultants)']

interior_leased = ['Outer Personnel (Professionals)',
                'Outer Personnel (Pain Share)',
                'Outer Personnel (Gain Share)']

resource_types = ['Resource change in',
                'Resource change out']

coverage_types = ['Restart because of change from Leased to Full',
                'Restart because of change from Leased to Temporary',
                'Restart because of change from Leased to Seasonal',
                'Restart because of change from Full to Leased',
                'Restart because of change from Temporary to Leased',
                'Restart because of change from Seasonal to Leased',
                'Ending because of change from Leased to Full',
                'Ending because of change from Leased to Temporary',
                'Ending because of change from Leased to Seasonal',
                'Ending because of change from Full to Leased',
                'Ending because of change from Temporary to Leased',
                'Ending because of change from Seasonal to Leased']

world_structure_changes = ['Continent Change',
                        'Country Change',
                        'City Change',
                        'State Change',
                        'Branch Change',
                        'Change below Branch']


world_types = ['Continent Transfer in',
            'Country Transfer in',
            'City Transfer in',
            'State Transfer in',
            'Branch Transfer in',
            'Transfer in below Branch',
            'Continent Transfer out',
            'Country Transfer out',
            'City Transfer out',
            'State Transfer out',
            'Branch Transfer out',
            'Transfer out below Branch']

#dictionaries with Movement Type names for positive & negative moves
positive_moves = {'Full to Temporary': 'Restart because of change from Full to Temporary',
                'Full to Leased': 'Restart because of change from Full to Leased',
                'Temporary to Full': 'Restart because of change from Temporary to Full',
                'Temporary to Leased': 'Restart because of change from Temporary to Leased',
                'Leased to Full': 'Restart because of change from Leased to Full',
                'Leased to Temporary': 'Restart because of change from Leased to Temporary',
                'Leased to Seasonal': 'Restart because of change from Leased to Seasonal',
                'Seasonal to Full': 'Restart because of change from Seasonal to Full',
                'Seasonal to Leased': 'Restart because of change from Seasonal to Leased',
                'Seasonal to Temporary': 'Restart because of change from Seasonal to Temporary',
                'Continent Change': 'Continent Transfer in',
                'Country Change': 'Country Transfer in',
                'City Change': 'City Transfer in',
                'State Change': 'State Transfer in',
                'Branch Change': 'Branch Transfer in',
                'Change below Branch': 'Transfer in below Branch',
                'Resource Type Change': 'Resource change in',
                'FTE Change': 'FTE Change in',
                'Other Movement Type': 'Other Movement Type: DC, Office, Level, Company, Program, Agreement',
                'Other TBC': 'Start'
                }

negative_moves = {'Full to Temporary': 'Ending because of change from Full to Temporary',
                'Full to Leased': 'Ending because of change from Full to Leased',
                'Temporary to Full': 'Ending because of change from Temporary to Full',
                'Temporary to Leased': 'Ending because of change from Temporary to Leased',
                'Leased to Full': 'Ending because of change from Leased to Full',
                'Leased to Temporary': 'Ending because of change from Leased to Temporary',
                'Leased to Seasonal': 'Ending because of change from Leased to Seasonal',
                'Seasonal to Full': 'Ending because of change from Seasonal to Full',
                'Seasonal to Leased': 'Ending because of change from Seasonal to Leased',
                'Seasonal to Temporary': 'Ending because of change from Seasonal to Temporary',
                'Continent Change': 'Continent Transfer out',
                'Country Change': 'Country Transfer out',
                'City Change': 'City Transfer out',
                'State Change': 'State Transfer out',
                'Branch Change': 'Branch Transfer out',
                'Change below Branch': 'Transfer out below Branch',
                'Resource Type Change': 'Resource change out',
                'FTE Change': 'FTE Change out',
                'Other Movement Type': 'Other Movement Type: DC, Office, Level, Company, Program, Agreement',
                'Other TBC': 'Ending'
                }

#dictionaries created to change Coverage Type if needed
cov_change = {'(interior coverage)' : '(exterior coverage)',
            '(exterior coverage)' : '(interior coverage)'}

coverage_change_dict = {'interior coverage' : 'exterior coverage',
            'exterior coverage' : 'interior coverage'}


def read_files(files, dtype_dict, current_date):
    """Takes all CSV files from files folder,
    turns them into DataFrames and concatenates them.
    Generates  the list of all months starting from the selected month.
    """
    list_of_dfs = []
    encoding_list = ['skdjsd', None, 'UTF-8', 'wwwd','GBK', 'iso-8859-1', 'latin_1']
    #list all the files in the files folder
    all_files = os.listdir(files)
    #reduce the list to csv files. Skip Temporaryorary hidden files.
    all_files = [x for x in all_files if 'csv' in x and x[0] != '.' and x[0] != '~']
    
    for mf in all_files:
        print(files+mf)
        for encoding in encoding_list:
            try:
                myfile = pd.read_csv(files+mf, dtype=dtype_dict, usecols=columns, encoding=encoding)
                df = myfile.copy()
                df = df.loc[df.ID.notna()]
                df.Date = pd.to_datetime(df.Date, format='%d/%m/%Y')
                df.GON = df.GON.str.zfill(8)
                df.Date = (df.Date.dt.floor('d') + pd.offsets.MonthEnd(0) - pd.offsets.MonthBegin(1))
                df.Date = pd.to_datetime(df['Date'])
                list_of_dfs.append(df)
            except: 
                print("Encoding", encoding, "is incorrect. Trying another one...")
            else:
                break
    df = pd.concat(list_of_dfs, ignore_index=True)

    if len(df[df['ID'].duplicated()]) > 0:
        duplicated_ids = df[df['ID'].duplicated()]['ID'].tolist()
        print("Duplicated IDs found", duplicated_ids, "\nThe program has ended.")
        sys.exit()

    all_months = df.loc[(df['Date'].dt.date >= current_date)].sort_values(by='Date').Date.dt.date.unique()
    return df, all_months


def date_range_reporting(df, current_date, prev_month):#TO CORRECT
    """Labels each movement as: Current Month/Previous Month/Old Month & adds Coverage Type column"""
    df['DateRange'] = np.where((df.Date.dt.year == current_date.year) & (df.Date.dt.month == current_date.month),
         "Current Month", np.nan)
    df['DateRange'] = np.where((df.Date.dt.year == prev_month.year) & (df.Date.dt.month == prev_month.month),
         "Previous Month", df.DateRange)
    df['DateRange'] = np.where(df.Date.dt.date < prev_month, "Old Month", df.DateRange)

    #add Coverage Type column based on values in Type of WTE and Resource Type columns ###CHANGE
    df['Coverage Type'] = np.where((df['Type of WTE'] == "Leased") & (df['Resource Type'].isin(interior_leased)),
                                    "(interior coverage)", "(exterior coverage)")
    df['Coverage Type'] = np.where((df['Type of WTE'] == "Seasonal"),
                                    "(interior coverage)", df['Coverage Type'])
    
    #limit DF to those GONs which have movements for start/current month and exclude future movements ###CHANGE
    current_gons = df.loc[df.DateRange == 'Current Month']['GON'].unique().tolist()
    df = df.loc[(df.GON.isin(current_gons)) & (df.DateRange != 'nan')]

    #create dictionary - key:ID, value:Coverage Type ###CHANGE
    coverage_dict = dict(zip(df['ID'], df['Coverage Type']))

    #create dictionary - key:ID, value:Type of WTE 
    wte_type_dict = dict(zip(df['ID'], df['Type of WTE']))

    #create dictionary - key:ID, value:WTE
    wte_dict = dict(zip(df['ID'], df['WTE']))

    #create dictionary - key:ID, value:Resource Type
    resource_type_dict = dict(zip(df['ID'], df['Resource Type']))

    return df, coverage_dict, wte_type_dict, wte_dict, resource_type_dict


def get_first_types(df):
    """Finds movements that should be marked as Hire/Ending/Restart/Other"""
    #Level each date to find the closest previous
    df['DateLevel'] = df.groupby('GON')['Date'].rank(method='dense', ascending=False)
    #GONs with positive moves in latest previous month
    previous_plus = df.loc[(df.DateLevel == 2) & (df['WTE'] > 0)]['GON'].unique().tolist()
    #GONs with old+previous months
    previous_moves = df.loc[df.DateRange.isin(['Old Month', 'Previous Month'])]['GON'].unique().tolist()
    #GONs with previous month moves
    prevmonth_moves = df.loc[df.DateRange == 'Previous Month']['GON'].unique().tolist()
    #GONs with previous moves in previous months
    pos_prevmonth_moves = df.loc[(df.DateRange == 'Previous Month') & (df['WTE'] > 0)]['GON'].unique().tolist()
    #GONs with positive moves in current monthfillna
    pos_curmonth_moves = df.loc[(df.DateRange == 'Current Month') & (df['WTE'] > 0)]['GON'].unique().tolist()
    #GONs with negative moves in current month
    neg_curmonth_moves = df.loc[(df.DateRange == 'Current Month') & (df['WTE'] < 0)]['GON'].unique().tolist()
    
    #additional columns
    df['Previous_Moves'] = np.where(df['GON'].isin(previous_moves), 1, 0)
    df['Prev_Month_Moves'] = np.where(df['GON'].isin(prevmonth_moves), 1, 0)
    df['Prev_Month_Plus'] = np.where(df['GON'].isin(pos_prevmonth_moves), 1, 0)
    df['Current_Month_Plus'] = np.where(df['GON'].isin(pos_curmonth_moves), 1, 0)
    df['Current_Month_Minus'] = np.where(df['GON'].isin(neg_curmonth_moves), 1, 0)
    df['Previous_Plus'] = np.where(df['GON'].isin(previous_plus), 1, 0)
    
    #Final Movement Types for part one:
    #Hire
    df['Movement Type'] = np.where((df.Previous_Moves == 0) & (df.Current_Month_Minus == 0) & (df.Current_Month_Plus == 1), "Start", np.nan)
    #Ending
    df['Movement Type'] = np.where((df.Current_Month_Plus == 0) & (df.Current_Month_Minus == 1), "Ending", df['Movement Type'])
    #Restart
    df['Movement Type'] = np.where((df.Previous_Moves == 1) & (df.Prev_Month_Plus == 0) & (df.Current_Month_Plus == 1) & (df.Current_Month_Minus == 0), "Restart", df['Movement Type'])
    #Other: (plus and minus in current month)
    df['Movement Type'] = np.where((df.Current_Month_Plus == 1) & (df.Current_Month_Minus == 1), "Other", df['Movement Type'])

    #if neither of the above Movement Types is assigned and this is TBC plus, mark it as 'Start'
    df['Movement Type'] = np.where((df['Movement Type'] == 'nan') & (df.GON.str.contains('TBC') & (df['WTE'] > 0)), "Start", df['Movement Type'])
    #if neither of the above Movement Types is assigned and this is TBC minus, mark it as 'Ending'
    df['Movement Type'] = np.where((df['Movement Type'] == 'nan') & (df.GON.str.contains('TBC') & (df['WTE'] < 0)), "Ending", df['Movement Type'])

    #identify those GONs which can replace Errors2
    potential_others = df.loc[df.Previous_Plus == 1]['GON'].unique().tolist()
    return df, potential_others


def generate_part_one(df):
    """Puts the first Movement Types in the target table"""
    df = df.loc[(df.DateRange == 'Current Month') & (df['Movement Type'] != 'Other')][['ID', 'GON', 'Movement Type']]
    df['Movement Type'] = df['Movement Type'].str.replace('nan', 'Error_1')
    return df


def get_second_types(df, potential_others, world_structure_changes):
    """Creates the remaining Movement Types"""
    #columns required
    cols = ['GON', 'ID', 'Continent', 'Country', 'City', 'State', 'Branch',
            'Spec_PDG_PC', 'Other_Moves', 'Resource Type', 'Type of WTE', 'WTE']
    
    #limit the DF to current month and Movement Type Other
    df = df.loc[(df.DateRange == 'Current Month') & (df['Movement Type'] == 'Other')]
    df = df.fillna('NNN')
    
    #concatenate columns
    df['Spec_PDG_PC'] = df['Specialization'] + df['PDG'] + df['Payment Center']
    df['Other_Moves'] = df['Office'] + df['Level'] + df['Dev Center'] + df['Company'] + df['Program'] + df['Agreement']
    
    #limit the DF to columns required
    df = df[cols]
    
    #count the number of unique GONs
    df['Moves_Count'] = df.groupby('GON')['GON'].transform('count')
    
    #if any real GONs (not tbcs) occur more than twice, delete the rows with them
    if len(df.loc[(df.Moves_Count > 2) & (~df.GON.str.contains('TBC'))]['GON'].unique()) > 0:
        print("\nGONs with more than two movements within the same month:",
              df.loc[(df.Moves_Count > 2) & (~df.GON.str.contains('TBC'))]['GON'].unique(), "\n")
        invalid_gons= df.loc[(df.Moves_Count > 2) & (~df.GON.str.contains('TBC'))]['GON'].unique().tolist()
        #print(invalid_GONs)
        invalid_list.extend(invalid_gons)
        df = df.loc[(df.Moves_Count < 3) | (df.GON.str.contains('TBC'))]

    #create DF for tbcs with more than two moves
    tbcs = df.loc[(df.Moves_Count > 2) & (df.GON.str.contains('TBC'))].copy()
    #if WTE is plus, mark as Start, else: Ending
    tbcs['Movement Type'] = np.where(tbcs['WTE'] > 0, "Start", "Ending")
    tbcs = tbcs[['ID', 'GON', 'Movement Type']]

    #now exclude those tbcs from main DF
    df = df.loc[df.Moves_Count < 3]
    
    #create two dfs: 1 - for postivite moves, 2 - for negative ones
    pos_df = df.loc[df['WTE'] > 0]
    neg_df = df.loc[df['WTE'] < 0]
    df = pd.merge(pos_df,neg_df, on='GON')
    
    ##Generate Movement Types for each GON
    #Other Movement Type
    df['Type'] = np.where((df['Other_Moves_y'] != df['Other_Moves_x']), "Other Movement Type", np.nan)
    #WTE Change
    df['Type'] = np.where((df['WTE_y'] + df['WTE_x'] != 0), "FTE Change", df['Type'])
    #Resource Type Change
    df['Type'] = np.where((df['Resource Type_y'] != df['Resource Type_x']),
                           "Resource Type Change", df['Type'])
    #Change below Branch
    df['Type'] = np.where((df['Spec_PDG_PC_y'] != df['Spec_PDG_PC_x']),
                           "Change below Branch", df['Type'])
    #Branch Change
    df['Type'] = np.where((df['Branch_y'] != df['Branch_x']), "Branch Change", df['Type'])
    #City Change
    df['Type'] = np.where((df['City_y'] != df['City_x']), "City Change", df['Type'])
    
    #State Change
    df['Type'] = np.where((df['State_y'] != df['State_x']) & (df['Country_y'].str.contains('USA')) & (df['Country_x'].str.contains('USA')), 
                          "State Change", df['Type'])
    #Country Change
    df['Type'] = np.where((df['Country_y'] != df['Country_x']), "Country Change", df['Type'])

    #Continent Change
    df['Type'] = np.where((df['Continent_y'] != df['Continent_x']), "Continent Change", df['Type'])

    #create a dictionary for GONs that belong to world_structure_changes key:GON, value:Type
    world_changes_dict = df[df['Type'].isin(world_structure_changes)].set_index('GON').to_dict()['Type']

    #Seasonal to Full
    df['Type'] = np.where((df['Type of WTE_y'] == 'Seasonal') & (df['Type of WTE_x'] == 'Full'),
                          "Seasonal to Full", df['Type'])
    #Full to Temporary
    df['Type'] = np.where((df['Type of WTE_y'] == 'Full') & (df['Type of WTE_x'] == 'Temporary'),
                          "Full to Temporary", df['Type'])
    #Full to Leased
    df['Type'] = np.where((df['Type of WTE_y'] == 'Full') & (df['Type of WTE_x'] == 'Leased'),
                          "Full to Leased", df['Type'])
    #Leased to Full
    df['Type'] = np.where((df['Type of WTE_y'] == 'Leased') & (df['Type of WTE_x'] == 'Full'),
                          "Leased to Full", df['Type'])
    #Temporary to Full
    df['Type'] = np.where((df['Type of WTE_y'] == 'Temporary') & (df['Type of WTE_x'] == 'Full'),
                          "Temporary to Full", df['Type'])
    #Leased to Temporary
    df['Type'] = np.where((df['Type of WTE_y'] == 'Leased') & (df['Type of WTE_x'] == 'Temporary'),
                          "Leased to Temporary", df['Type'])
    #Temporary to Leased
    df['Type'] = np.where((df['Type of WTE_y'] == 'Temporary') & (df['Type of WTE_x'] == 'Leased'),
                          "Temporary to Leased", df['Type'])
    #Seasonal to Leased
    df['Type'] = np.where((df['Type of WTE_y'] == 'Seasonal') & (df['Type of WTE_x'] == 'Leased'),
                          "Seasonal to Leased", df['Type'])
    #Leased to Seasonal
    df['Type'] = np.where((df['Type of WTE_y'] == 'Leased') & (df['Type of WTE_x'] == 'Seasonal'),
                          "Leased to Seasonal", df['Type'])
    #Seasonal to Temporary
    df['Type'] = np.where((df['Type of WTE_y'] == 'Seasonal') & (df['Type of WTE_x'] == 'Temporary'),
                          "Seasonal to Temporary", df['Type'])
    #Other TBC: if Type is nan and GON is TBC
    df['Type'] = np.where((df['Type'] == 'nan') & (df.GON.str.contains('TBC')),
                          "Other TBC", df['Type'])
    #Other Movement Type if type is nan and GON is in potential_others
    df['Type'] = np.where((df['Type'] == 'nan') & (df.GON.isin(potential_others)),
                          "Other Movement Type", df['Type'])
    return df, tbcs, world_changes_dict


def generate_part_two(df, df_tbcs, pos=positive_moves, neg=negative_moves):
    """Creates the next Movement Types"""
    #map Final Movement Types
    df['Type_plus'] = df['Type'].map(pos)
    df['Type_minus'] = df['Type'].map(neg)
    
    #limit the dfs to 3 columns and rename them
    positive_ids = df[['ID_x', 'GON', 'Type_plus']].rename(columns={'ID_x': 'ID', 'Type_plus': 'Movement Type'})
    negative_ids = df[['ID_y', 'GON', 'Type_minus']].rename(columns={'ID_y': 'ID', 'Type_minus': 'Movement Type'})
    
    #combine three dataframes & drop duplicates
    p2 = pd.concat([positive_ids, negative_ids, df_tbcs], ignore_index=False)
    p2 = p2.drop_duplicates()
    
    #replace NaNs with Error 2
    p2 = p2.fillna("Error_2")
    p2['Movement Type'] = p2['Movement Type'].str.replace('nan', 'Error 2')
    return p2


def concatenate_dfs(df1, df2):
    """Concatenates two parts and add start month"""
    df = pd.concat([df1, df2], ignore_index=False)
    df['Start Date'] = start_month
    return df


def final_action(df, coverage_types, cov_dict, world_changes_dict, wte_type_dict, cov_change, resource_type_dict):
    """Adds Coverage Type where necessary"""
    #get Coverage Type for each id
    df['cov_type'] = df['ID'].map(cov_dict)
    #show number of unique Coverage Types per GON
    df['cov_type_count'] = df.groupby('GON')['cov_type'].transform('nunique')
    #get Type of WTE per each id
    df['wte_type'] = df['ID'].map(wte_type_dict)
    #get WTE per each id
    df['WTE'] = df['ID'].map(wte_dict)

    df['Resource Type'] = df['ID'].map(resource_type_dict)
    df['res_type_count'] = df.groupby('GON')['Resource Type'].transform('nunique')

    #show number of unique Type of WTEs per GON
    df['wte_type_count'] = df.groupby('GON')['wte_type'].transform('nunique')
    #create list of GONs that have Leased Type of WTE
    gons_leased = df[df['wte_type'] == 'Leased']['GON'].unique().tolist()
    #show if GON has Leased Type of WTE
    df['Is_Leased'] = np.where(df['GON'].isin(gons_leased), 1, 0)

    #create two additional columns and add specific values to them
    df['From'] = np.nan
    df['To'] = np.nan

    df['From'] = np.where((df['cov_type_count'] == 1), df['cov_type'].str[1:-1], df['From'])
    df['From'] = np.where((df['cov_type_count'] == 2) &
                          (df['WTE'] < 0),
                          df['cov_type'].str[1:-1], df['From'])
    df['From'] = np.where((df['cov_type_count'] == 2) &
                          (df['WTE'] > 0),
                          df['cov_type'].str[1:-1].map(coverage_change_dict), df['From'])

    df['To'] = np.where((df['cov_type_count'] == 1), df['From'], df['To'])
    df['To'] = np.where((df['cov_type_count'] == 2) &
                        (df['WTE'] < 0),
                        df['cov_type'].str[1:-1].map(coverage_change_dict), df['To'])
    df['To'] = np.where((df['cov_type_count'] == 2) &
                        (df['WTE'] > 0),
                        df['cov_type'].str[1:-1], df['To'])

    #change Coverage Type for the one whch is assigned to the 'Leased ID' if per GON:
    #there are two different Coverage Types & Type of WTEs and one of the IDs is Leased
    df['cov_type'] = np.where((df['cov_type_count'] == 2) &
                              (df['wte_type_count'] == 2) &
                              (df['Is_Leased'] == 1) &  
                              (df['wte_type'] != 'Leased'), df['cov_type'].map(cov_change), df['cov_type'])

    
    #if Movement Type belongs to coverage_types and contains Leased, add required Coverage Type
    df['Final Movement Type'] = np.where((df['Movement Type'].isin(coverage_types)) & (df['Movement Type'].str.contains('Leased')),
                               df.apply(lambda x: x['Movement Type'].replace('Leased', 'Leased '+x['cov_type']), axis=1),
                               np.nan)
    
    #if Movement Type is in coverage_types and Movement Types str does not contain Leased, add Coverage Type at the end
    df['Final Movement Type'] = np.where((df['Movement Type'].isin(coverage_types)) & (~df['Movement Type'].str.contains('Leased')),
                               df['Movement Type'] + ' ' + df['cov_type'], df['Final Movement Type'])

    #add world change column and map values to it
    df['world_change'] = df['GON'].map(world_changes_dict)

    #create final movement types:
    df['Final Movement Type'] = np.where((df['Final Movement Type'].str.contains('|'.join(['Ending because', 'Restart because']))) &
                               (df['world_change'].notna()) &
                               (df['Is_Leased'] == 0),
                               df['Movement Type'] + ' - ' + df['world_change'], df['Final Movement Type'])
    
    df['Final Movement Type'] = np.where((df['Final Movement Type'].str.contains('|'.join(['Ending because', 'Restart because']))) &
                               (df['world_change'].notna()) &
                               (df['Is_Leased'] == 1),
                               df['Movement Type'] + ' - ' + df['world_change'], df['Final Movement Type'])

    df['Final Movement Type'] = np.where((df['Movement Type'].isin(resource_types)),
                                     df['Movement Type'] + ' (' + df['From'] + ' to ' + df['To'] + ')',
                                     df['Final Movement Type'])

    df['Final Movement Type'] = np.where((df['Movement Type'].isin(world_types)) &
                              (df['Is_Leased'] == 1) &
                              (df['res_type_count'] == 2) &
                              (df['wte_type_count'] == 1),
                                df['Movement Type'] + ' (Leased - ' + df['From'] + ' to ' + df['To'] + ')',
                                df['Final Movement Type'])
    
    df['Final Movement Type'] = np.where((df['Final Movement Type'].isna()),
                                     (df['Movement Type']), df['Final Movement Type'])

    #replace Movement Type with Final Movement Type                         
    df['Movement Type'] = df['Final Movement Type']

    #sort values by GON
    df = df.sort_values(by='GON')
    
    #drop unnecessary columns
    df = df.iloc[:,:4]
    return df


def df_to_excel(df):
    """Saves the final file as xlsx with current date and time"""
    now = datetime.now()
    date_now = now.strftime('%d-%b-%Y_%H-%M-%S')
    df.to_excel('results/mov_types_'+date_now+'.xlsx', index=False)


if __name__ == '__main__':
    from datetime import datetime

    #folder with files to read
    files = 'files/'
    #files for movements with more than two moves per month
    more_moves = "GONs_two_plus.txt"

    #enter the start date
    enter_date = input("\nENTER THE START DATE in the format: [month-year]: ")

    start_time = datetime.now()

    try:
        current_date = datetime.strptime(enter_date, '%m-%Y').date()
    except ValueError:
        print("Incorrect date format")
        sys.exit()

    print("\nWork in progress..")

    #concatenate all csv files and identify all_months in them
    big_file, all_months = read_files(files, dtype_dict, current_date)

    #create a list for invalid GONs
    invalid_list = []

    #clear elements_not_found.txt
    open(more_moves, "w").close()

    #crate a list for dataframes
    dfs_to_concat = []

    print("\n")

    for m in all_months:
        print(m)
        start_month = m
        prev_month = start_month+dateutil.relativedelta.relativedelta(months=-1) 
        daterange, cov_dict, wte_type_dict, wte_dict, resource_type_dict = date_range_reporting(big_file.copy(), start_month, prev_month)
        first_types, potential_others = get_first_types(daterange.copy())
        part_one = generate_part_one(first_types.copy())
        second_types, tbcs, world_changes_dict = get_second_types(first_types.copy(), potential_others, world_structure_changes)
        part_two = generate_part_two(second_types.copy(), tbcs.copy())
        part_three = concatenate_dfs(part_one, part_two)
        part_four = final_action(part_three, coverage_types, cov_dict, world_changes_dict, wte_type_dict, cov_change, resource_type_dict)
        dfs_to_concat.append(part_four)

    new_df = pd.concat(dfs_to_concat, ignore_index=True)
    column_names = ['Start Date', 'ID', 'GON', 'Movement Type']
    new_df = new_df[column_names]

    df_to_excel(new_df)

    with open(more_moves, 'w') as f:
        for line in invalid_list:
            f.write(f"{line}\n")
    
    print("final number of rows:", len(new_df))

    end_time = datetime.now()
    print("\nJob complete. Duration: {}".format(end_time - start_time))