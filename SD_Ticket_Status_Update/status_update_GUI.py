# Author: jakub-kuba

from tkinter import *
import tkinter as tk
from tkinter import messagebox as msb
from tkinter.filedialog import askopenfile
import pandas as pd
import numpy as np
from datetime import datetime
import os

class Application(Frame):
    """App for update statuses"""
    def __init__(self, master):
        """Initialises Frame"""
        super(Application, self).__init__(master)
        self.grid()
        self.create_widgets()

    def create_widgets(self):
        self.txt_edit = tk.Text(root, background="floral white",
                                fg="gray38", font="sans 12 bold",
                                state="disabled")
        fr_buttons = tk.Frame(root, background="gray38", bd=8)
        
        self.btn_upload = tk.Button(fr_buttons, bg="dark green",
                                  fg="floral white", text="START",
                                  activebackground="gray38",
                                  activeforeground="floral white",
                                  command = self.action,
                                  width=12, height=4,
                                  font="sans 16 bold")
        
        self.btn_upload.grid(row=0, column=0, sticky="ew",
                             padx=16, pady=20)
        
        fr_buttons.grid(row=0, column=0, sticky="ns")
        self.txt_edit.grid(row=0, column=1, sticky="nsew")

        message1 = "\n\nUpload the files:\n"
        message2 = "\n  1. Report - ACTIVE (csv)"
        message2 += "\n  2. Report - COMPLETED (csv)"
        message2 += "\n  3. SYSTEM Report (xlsx)\n"

        self.text_update(message1, message2)

    def action(self):
        """Executes the code after pressing the start button"""

        def status_update(mylist,mydf,colname):
            """Gets the full status of SD ID"""
            for x in mylist:
                try:
                    mydf[colname] = np.where((mydf[x] > 0), x, mydf[colname])
                except KeyError:
                    continue
            return mydf[colname]
        
        def rejected_to_completed(df_active,df_completed):
            """Moves 'Rejected' rows from 'Active' to 'Completed'"""
            cl_columns = df_completed.columns.to_list()
            rejected = df_active.loc[df_active['SD Status'] == 'Rejected'][cl_columns]
            new_df = pd.merge(df_completed, rejected, how='outer')
            return new_df

        def manipulate(df_report, kind, sys_ids, status, drop=True):
            """Manipulates DataFrame"""
            rep = df_report.copy()
            rep.rename(columns={"System ID": "id",
                            "System Reference": "second_id",
                            "SD ID": "SD Ticket ID"}, inplace=True)

            if kind == "actives":
                #remove 'Rejected' rows
                rep = rep.loc[rep['SD Status'] != 'Rejected']
                #insert new column 'status'
                rep['status'] = 'website visit'
                rep['status'] = status_update(mapp, rep, 'status')
                #insert new column 'category'
                rep['category'] = rep['status'].map(status_dict)
                #insert new column 'SD Ticket Status'
                rep['SD Ticket Status'] = np.where((rep['SD Status'].isin(['Active', 'Suspended'])),
                                        rep['SD Status']+" - "+rep['category'],
                                        rep['SD Status'])
            
            #drop rows with empty id columns
            rep.dropna(subset=['id', 'second_id'], how='all', inplace=True)

            #operations on 'id' & 'second_id'
            rep['new_id'] = np.where((~rep['id'].str.contains("-", na=False)),
                                    rep['id']+"-1", rep['id'])
            rep['new_second_id'] = np.where((~rep['second_id'].str.contains("-", na=False)),
                                    rep['second_id']+"-1", rep['second_id'])
            rep['new_id'].fillna(rep['new_second_id'], inplace=True)
            rep['no_new_id'] = rep['new_id'].str[-1:]
            rep['new_id_num'] = pd.to_numeric(rep['new_id'].str[:-2], errors='coerce').astype('Int64')
            rep = rep.loc[(rep['new_id_num'].notna()) & (rep['new_id_num'].between(min_system, max_system))]

            #get assistant in the correct format
            rep['Assistant'] = rep['Assistant'].str.replace("  ", " ")
            rep['Assistant'] = rep['Assistant'].str.strip()
            rep['Assistant'] = np.where((rep['Assistant'].str.count("\(") > 1),
                                            rep['Assistant'].str[:-11],
                                            rep['Assistant'])
            assistant = rep['Assistant'].str.split(',', expand=True)
            rep['Assistant'] = assistant[1].str[1:-11]+" "+assistant[0]+assistant[1].str[-11:]


            if drop:
                rep = rep.drop_duplicates(subset=['new_id'], keep=False)
            if kind == 'actives':
                rep = rep[['new_id_num', 'no_new_id', 'SD Ticket ID', 'SD Ticket Status', 'Assistant']]
                rep.rename(columns={"new_id_num": "System Request ID"}, inplace=True)
            else:
                rep = rep[['new_id_num', 'no_new_id', 'SD Ticket ID', 'SD Status', 'Assistant']]
                rep.rename(columns={"new_id_num": "System Request ID",
                                    "SD Status": "SD Ticket Status"}, inplace=True)
            if status == 'accepted':
                sys_ids = sys_ids.loc[(sys_ids['Status'].isin(['Request Accepted'])) & (sys_ids['Access Type'].isin(['Full', 'Standard']))][all_columns]
            else:
                sys_ids = sys_ids.loc[sys_ids['Access Type'].isin(['Full', 'Standard'])][all_columns]

                rep = pd.merge(rep, sys_ids, how='left')
                rep = rep.loc[(rep['SD Ticket ID'] == rep['SD Ticket ID 1']) |
                            (rep['SD Ticket ID'] == rep['SD Ticket ID 2']) |
                            (rep['SD Ticket ID'] == rep['SD Ticket ID 3']) |
                            (rep['SD Ticket ID'] == rep['SD Ticket ID 4']) |
                            (rep['SD Ticket ID'] == rep['SD Ticket ID 5'])
                ]
                rep = rep.iloc[:,:5]

            sys_ids['System Request ID'] = pd.to_numeric(sys_ids['System Request ID'], errors='coerce').astype('Int64')
            sys_ids = sys_ids.loc[sys_ids['System Request ID'].notna()]

            #remove potential duplicates
            rep = rep.drop_duplicates(subset=['System Request ID', 'no_new_id'], keep=False)

            rep = rep.set_index(['System Request ID', 'no_new_id']).unstack(level=1).sort_index(axis=1,level=1)
            rep.columns = rep.columns.map(' '.join)
            rep = rep.reset_index()
            columns_now = rep.columns.to_list()
            missing_cols = list(set(all_columns) - set(columns_now))
            rep[missing_cols] = ""
            rep = rep[all_columns]

            rep = pd.merge(rep, sys_ids, how='inner', on='System Request ID')
            rep.fillna('', inplace=True)

            if status == 'accepted':
                rep = rep.loc[(rep['SD Ticket ID 1_x'] != rep['SD Ticket ID 1_y']) |
                            (rep['SD Ticket ID 2_x'] != rep['SD Ticket ID 2_y']) |
                            (rep['SD Ticket ID 3_x'] != rep['SD Ticket ID 3_y']) |
                            (rep['SD Ticket ID 4_x'] != rep['SD Ticket ID 4_y']) |
                            (rep['SD Ticket ID 5_x'] != rep['SD Ticket ID 5_y'])
                ]
            else:
                rep = rep.loc[(rep['SD Ticket Status 1_x'] != rep['SD Ticket Status 1_y']) |
                            (rep['SD Ticket Status 2_x'] != rep['SD Ticket Status 2_y']) |
                            (rep['SD Ticket Status 3_x'] != rep['SD Ticket Status 3_y']) |
                            (rep['SD Ticket Status 4_x'] != rep['SD Ticket Status 4_y']) |
                            (rep['SD Ticket Status 5_x'] != rep['SD Ticket Status 5_y'])
                ]

            rep = rep.replace('', np.nan, regex=True)
            xs = rep.columns.to_list()[1:16]
            ys = rep.columns.to_list()[16:]
            for x, y in zip(xs, ys):
                rep[x].fillna(rep[y], inplace=True)
            rep = rep.iloc[:,:16]
            rep.columns = rep.columns.str.replace('_x', '')

            return rep

        def final_report(new_reports, act_status, com_status, sys_ids):
            """Creates final files"""
            sys_rep = sys_ids.copy()
            sys_rep = sys_rep.loc[sys_rep['Access Type'].isin(['Full', 'Standard'])][all_columns]
            sys_rep['System Request ID'] = pd.to_numeric(sys_rep['System Request ID'], errors='coerce').astype('Int64')
            sys_rep = sys_rep.loc[sys_rep['System Request ID'].notna()]
            sys_rep.iloc[:,1:] = sys_rep.iloc[:,1:].astype(object)

            full_status = pd.concat([act_status, com_status], axis=0)
            full = pd.concat([new_reports, full_status], axis=0)
            full = full.drop_duplicates(subset=['System Request ID'])
            full.iloc[:,1:] = full.iloc[:,1:].astype(object)
            #delete those rows from 'full' which are already in 'sys_ids'
            full = pd.merge(full, sys_rep, indicator=True, how='outer').query('_merge=="left_only"').drop('_merge', axis=1)

            now = datetime.now()
            date = now.strftime('%d-%b-%Y')

            len_full = len(full)
            div = len(full) / row_limit
            mult = int(div)
            sub = len_full - (mult * row_limit)

            #create 'results' folder if doesn't exist
            path = 'results'
            isExist = os.path.exists(path)

            if not isExist:
                os.makedirs(path)

            #if number of rows is more than row_limit, split the file into parts
            if len_full <= row_limit:
                full.to_excel("results/ticket_sd_status "+ date+" ({:02}).xlsx".format(1), sheet_name='SystemReport', index=False)
            elif sub == 0:
                i=1
                for huge_df in np.array_split(full, len_full // row_limit):
                    huge_df.to_excel("results/ticket_sd_status " +date+" ({:02}).xlsx".format(i), sheet_name='SystemReport', index=False)
                    i += 1
            else:
                new_full = full[:-sub]
                rest = full[-sub:]
                i=1
                for huge_df in np.array_split(new_full, len(new_full) // row_limit):
                    huge_df.to_excel("results/ticket_sd_status " +date+" ({:02}).xlsx".format(i), sheet_name='SystemReport', index=False)
                    i += 1
                rest.to_excel("results/ticket_sd_status " +date+" ({:02}).xlsx".format(i), sheet_name='SystemReport', index=False)
            
            full.to_excel("results/Full_ticket_sd_status " +date+".xlsx", sheet_name='SystemReport', index=False)
        

        panel = "my panel.xlsx"
        #check if the panel file exists
        try:
            mapping = pd.read_excel(panel, sheet_name='mapping')
        except FileNotFoundError:
            message = "\n\n\n\n'My Panel' file not found!\n"
            self.text_update(message)
            msb.showerror(title="Error", message="my panel.xlsx not found!")
            return None
        
        #create a list of values of Status Name, except the first value
        mapp = [x for x in mapping['Status Name'] if str(x) != 'website visit']

        #take necessary values from panel
        status_dict = dict(zip(mapping['Status Name'], mapping['Category']))
        row_limit = int(mapping['Row Limit'][0])
        min_system = int(mapping['System ID Min'][0])
        max_system = int(mapping['System ID Max'][0])

        #create a list of necessary columns in 'System Report'
        all_columns = ['System Request ID', 'SD Ticket ID 1', 'SD Ticket Status 1', 'Assistant 1',
                'SD Ticket ID 2', 'SD Ticket Status 2', 'Assistant 2',
                'SD Ticket ID 3', 'SD Ticket Status 3', 'Assistant 3',
                'SD Ticket ID 4', 'SD Ticket Status 4', 'Assistant 4',
                'SD Ticket ID 5', 'SD Ticket Status 5', 'Assistant 5']

        message = "\n\n\n\nPlease wait...\n"
        self.text_update(message)

        file1 = askopenfile(title = "report active", mode = "r", filetypes = [("CSV Files", "*.csv")])
        file2 = askopenfile(title = "report completed", mode = "r", filetypes = [("CSV Files", "*.csv")])
        file3 = askopenfile(title = "System Report", mode = "r", filetypes = [("Excel Files", "*.xlsx")])

        if file1 and file2 and file2 is not None:
            try:
                active = pd.read_csv(file1.name, dtype={'System ID': str})
                completed = pd.read_csv(file2.name, dtype={'System ID': str})
                sys_ids = pd.read_excel(file3.name)
            except pd.errors.ParserError:
                message = "\n\n\nThe files must have CSV extension.\n"
                message += "\nResave the files and try again.\n"
                self.text_update(message)
                msb.showerror(title="Error", message="Make sure the fikes are saved as CSV!")
                return None
            file1.close()
            file2.close()
            file3.close()

            #move 'Rejected' rows from 'report active' to 'report completed'
            new_completed = rejected_to_completed(active,completed)
            #identify new SD Ticket IDs
            new_reports = manipulate(active, 'actives', sys_ids, 'accepted')
            #update SD Ticket Statuses based on 'report active'
            act_status =  manipulate(active, 'actives', sys_ids, 'full', drop=False)
            #update SD Ticket Statuses based on 'report completed'
            com_status =  manipulate(new_completed, 'completes', sys_ids, 'full', drop=False)
            #create xlsx files
            final_report(new_reports, act_status, com_status, sys_ids)

            message = "\n\n\n\nJob complete!\n"
            self.text_update(message)
            msb.showinfo(title="Info", message="Job complete!")

        else:
            message = "\n\n\nNot all the files have been imported!\n"
            message += "\nTry again.\n"
            self.text_update(message)

    def text_update(self, message, message2=""):
        """Edits the textbox"""
        self.txt_edit.tag_configure("center", justify="center")
        self.txt_edit.configure(state="normal")
        self.txt_edit.delete(0.0, END)
        self.txt_edit.insert(0.0, message2)
        self.txt_edit.insert(0.1, message, "center")
        self.txt_edit.configure(state="disabled")

root = tk.Tk()
root.title("SD Ticket Status Upate")
root.geometry("554x172")
root.resizable(False, False)
app = Application(root)
root.rowconfigure(0, minsize=75, weight=1)
root.columnconfigure(1, minsize=50, weight=3)
root.mainloop()