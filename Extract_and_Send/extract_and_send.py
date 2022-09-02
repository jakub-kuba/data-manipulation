# Author: jakub-kuba

import os, sys
from os import path
import pandas as pd
from datetime import datetime, date, timedelta
from openpyxl import load_workbook
import win32com.client


def open_excel(filename,sheets):
    """
    Check if: 1 - the necessary file exists,
    2 - it contains the required sheets,
    3 - and the sheets contain the required columns
    """
    for s in sheets:
        try:
            wb = load_workbook(filename,read_only=True)
            wb[s]
            pd.read_excel(filename,sheet_name=s,usecols=sheets[s])
        except FileNotFoundError:
            print(f"'{filename}' not found. The program has quit.\n")
            sys.exit()
        except KeyError:
            print(f"Worksheet: '{s}' not found in: '{filename}'. The program has quit.\n")
            sys.exit()
        except ValueError:
            print(f"Mandatory columns not found in: '{s}' worksheet in: '{filename}'. The program has quit.\n")
            sys.exit()


def prevday_lastmonth():
    """Get the last day of the previous month"""
    fs = date.today().replace(day=1)
    ls = (fs - timedelta(days=1)).strftime("%d-%B-%Y")
    return ls


def ask_display_send(question):
    """Choose 'display' or 'send' emails"""
    response = None
    while response not in ("d", "s"):
        response = input(question).lower()
    return response


def send_email(emails,unit,mail_message,file_name,destination,my_cycle,answer):
    """Sending emails"""
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = emails
    mail.CC = "xyz2972163zsd@gmail.com"
    mail.Subject = f'My system - {unit}: job complete {my_cycle}'
    br = '<br>'
    brss = '<br><br>'
    hello = 'Hi Everyone,'+ brss
    regards = 'Regards,'+br+'Our Team'
    mail.HTMLBody = hello+brss.join(mail_message)+brss+regards
    mail.Attachments.Add(destination+file_name)
    if answer == "d":
        mail.Display()
    else:
        mail.Send()


def main():
    print("\n")
    #display or send email
    answer = ask_display_send("Do you want emails to be displayed or sent immediately? (d - Display /s - Send): ")
    start_time = datetime.now()

    # define paths to folders
    files  = "C:/Users/JJ/Python/Extract_and_Send/source file/"
    destination  = "C:/Users/JJ/Python/Extract_and_Send/final files/"
    elems_not_found = "elements not found.txt"

    #check if the required folders exist
    if path.exists(files) == False or path.exists(destination) == False:
        print("The 'Source file' folder and/or the 'final files' do not exist. " + \
        "The program has quit.\n")
        sys.exit()

    #declare the panel file, its required sheets and columns
    panel = "control panel.xlsx"
    panel_sheetcols = {'cycles': ['Month','Cycle'],
                      'units_required': ['Unit Name','Emails','Email Type','Operation Type'],
                      'emails': ['email 1','email 2'],
                      'systems': ['Elems','System 1 Elems','System 2 Elems']}

    #check if the 'source file' folder contains only one xlsx file
    all_files = os.listdir(files)
    #exclude hidden temporary files
    xlsx_file = [x for x in all_files if x[-5:] == '.xlsx' and x[0] != '~']
    if len(xlsx_file) > 1:
        print("Too many xlsx files in the 'source file' folder! The program has ended\n")
        sys.exit()
    elif len(xlsx_file) < 1:
        print("The xlsx file not found in the 'source file' folder! The program has ended\n")
        sys.exit()  
    #get the location/name of the source file
    source_file = files+xlsx_file[0]

    #declare the required sheets and columns of the 'source file'
    source_sheetcols = {'current data': ['ID','Continent','Region','Country','ISO-alpha3 Code'],
                        'summary': ['Item','System 1 Region','System 2 Continent','System 1 Country']}

    print("Analysing:",panel,"\n")
    open_excel(panel,panel_sheetcols)
    print("Analysing:",xlsx_file[0],"\n")
    open_excel(source_file,source_sheetcols)

    ##turning the 'Control Panel' sheets into DataFrames
    cycles = pd.read_excel(panel, sheet_name = 'cycles')
    units = pd.read_excel(panel, sheet_name = 'units_required')
    mails = pd.read_excel(panel, sheet_name = 'emails')
    systems = pd.read_excel(panel, sheet_name = 'systems')

    ##create necessary dicts
    dict_cycles = dict(zip(cycles['Month'], cycles['Cycle']))
    sys_one_dict = dict(zip(systems['Elems'], systems['System 1 Elems']))
    sys_two_dict = dict(zip(systems['Elems'], systems['System 2 Elems']))
    email_dict = dict(zip(units['Unit Name'], units['Emails']))
    email_type_dict = dict(zip(units['Unit Name'], units['Email Type']))
    op_type_dict = dict(zip(units['Unit Name'], units['Operation Type']))

    ##create necessary lists
    unit_list = units['Unit Name'].to_list()
    sys_list = systems['Elems'].to_list()

    #get the text of two emails
    ls = prevday_lastmonth()
    email_cols = ['email 1', 'email 2']
    mails[email_cols] = mails[email_cols].replace('<last day of the previous month>', ls, regex=True)
    #!= 'nan'  prevents from adding empty cells to the list
    mail_text1 = [x for x in mails['email 1'].to_list() if str(x) != 'nan']
    mail_text2 = [x for x in mails['email 2'].to_list() if str(x) != 'nan']

    #loading the source file
    all_sheets =  pd.ExcelFile(source_file)
    current = pd.read_excel(all_sheets, 'current data',keep_default_na=False)
    summary = pd.read_excel(all_sheets, 'summary')

    #clear 'elements not found.txt'
    open(elems_not_found, "w").close()

    #get the cycle number
    this_month = datetime.now().month
    my_cycle = dict_cycles[this_month]

    print("Going through units:")
    for unit in unit_list:
        print(unit)
        elem_name = ""
        for s in sys_list:
            #look for the unit in Continent, Region and Country columns
            #str[:-8] is used to limit search to name only
            unit_len = len(current.loc[current[s].str[:-8].isin(unit.split(","))])
            if unit_len > 0:
                elem_name = s
                break
        if elem_name == "": #element not found: add its name to txt file and go to the next unit
            print(f"Unit: '{unit}' not found in the file")
            with open('elements not found.txt', 'a') as myfile:
                myfile.write(unit+"\n")
            continue
        
        #get the necessary information from control panel
        type_op = op_type_dict[unit]
        sys_one_elem = sys_one_dict[s]
        sys_two_elem = sys_two_dict[s]
        emails = email_dict[unit]
        type_of_email = email_type_dict[unit]

        if type_of_email == "email 1":
            mail_message = mail_text1   
        else:
            mail_message = mail_text2

        #create a name of the file
        file_name = unit+ " "+my_cycle+".xlsx"

        #filter 'Summary' for both type of operations
        #from summary sheet take units from System 1... or System 2... columns
        summ = summary.loc[(summary[sys_one_elem].str[:-8].isin(unit.split(",")) |
                            summary[sys_two_elem].str[:-8].isin(unit.split(",")))]

        #adddidtional sheets for 'full' operation                    
        if type_op != "simple":
            tbms = current[current['ID'].str.contains('TBM') & current[s].str[:-8].isin(unit.split(","))]
            curr = current.loc[current[s].str[:-8].isin(unit.split(","))]

        #save a new file with the required sheets
        with pd.ExcelWriter(destination+file_name) as writer:
            if type_op == "simple":
                summ.to_excel(writer, sheet_name='summary', index=False)
            else:
                curr.to_excel(writer, sheet_name='current data', index=False)
                tbms.to_excel(writer, sheet_name='TBMs', index=False)
                summ.to_excel(writer, sheet_name='summary', index=False)

        #generate email
        send_email(emails,unit,mail_message,file_name,destination,my_cycle,answer)

    end_time = datetime.now()
    print("\nJob complete. Duration: {}".format(end_time - start_time))


if __name__== "__main__":
    main()