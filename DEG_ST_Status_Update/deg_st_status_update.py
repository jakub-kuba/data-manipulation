# Author: jakub-kuba

from tkinter import *
import tkinter as tk
from tkinter import messagebox as msb
from tkinter.filedialog import askopenfile

import pandas as pd
import numpy as np
from datetime import datetime
import functools as ft

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
        def count_statuses(series, report):
            """Counts number of rows for selected group"""
            if report == 'active':
                return (series.isin(statuses_active)).sum()
            if report == 'completed':
                return (series.isin(statuses_completed)).sum()
            if report == 'canceled':
                return (series.isin(statuses_canceled)).sum()

        def status_update(mylist,mydf,colname):
            """Updates status"""
            for x in mylist:
                try:
                    mydf[colname] = np.where((mydf[x] > 0), x, mydf[colname])
                except KeyError:
                    continue
            return mydf[colname]

        def prepare_reports(degs, actives, completed, deg_status):
            """Creates degs_accepted, degs_fulls & big_rep DataFrames"""

            # create df app for add new serviceticket to degs (Accepted only)
            degs_accepted = degs.loc[(degs['Status'].isin(['Request Accepted'])) &
                                       (degs['Access Type'].isin(['Full', 'Trial']))][all_columns].copy()

            # convert deg it to integer
            degs_accepted['DEG System ID'] = pd.to_numeric(degs_accepted['DEG System ID'],
                                                                      errors='coerce').astype('Int64')

            # remove old deg ids (with prefix and the ones which are out of the required range (min_deg, max_deg))
            degs_accepted = degs_accepted.loc[degs_accepted['DEG System ID'].notna() &
                                                  (degs_accepted['DEG System ID'].between(min_deg, max_deg))]


            # create df app for new STs and status & Assistant update
            degs_fulls = degs.loc[(degs['Access Type'].isin(['Full', 'Trial']))][all_columns_plus].copy()

            # convert deg it to integer
            degs_fulls['DEG System ID'] = pd.to_numeric(degs_fulls['DEG System ID'],
                                                                 errors='coerce').astype('Int64')

            # remove old deg ids (with prefix and the ones which are out of the required range (min_deg, max_deg))
            degs_fulls = degs_fulls.loc[degs_fulls['DEG System ID'].notna() &
                                        (degs_fulls['DEG System ID'].between(min_deg, max_deg))]

            # concatenate actives and completed
            big_rep = pd.concat([actives, completed], axis=0)
            
            # variable created to be saved to excel
            actives_completed = big_rep.copy()

            # change name of columns
            big_rep = big_rep.rename(columns={"DEG System ID": "id",
                                              "Alternative No.": "ast_id",
                                              "ST Reference": "ServiceTicket ID"})

            # get new ServiceTicket Status based on the mapping table
            big_rep['status'] = 'website visit'
            big_rep['status'] = status_update(mapp, big_rep, 'status')
            big_rep['category'] = big_rep['status'].map(status_dict)
            big_rep['ServiceTicket Status'] = np.where((big_rep['Current Status'].isin(['Active', 'Suspended'])),
                                    big_rep['Current Status']+" - "+big_rep['category'],
                                    big_rep['Current Status'])

            # drop rows where there is no deg ID (id and ast_id)
            big_rep = big_rep.dropna(subset=['id', 'ast_id'], how='all')

            # put deg ID in one column ([deg_num] and BR location number in the other (['no_deg'])
            big_rep['deg_id'] = np.where((~big_rep['id'].str.contains("-", na=False)), big_rep['id']+"-1", big_rep['id'])
            big_rep['astp_id'] = np.where((~big_rep['ast_id'].str.contains("-", na=False)),
                                    big_rep['ast_id']+"-1", big_rep['ast_id'])
            big_rep['deg_id'] = big_rep['deg_id'].fillna(big_rep['astp_id'])
            big_rep['no_deg'] = big_rep['deg_id'].str[-1:]
            big_rep['deg_num'] = pd.to_numeric(big_rep['deg_id'].str[:-2], errors='coerce').astype('Int64')

            big_rep = big_rep.loc[(big_rep['deg_num'].notna()) & (big_rep['deg_num'].between(min_deg, max_deg))]

            #map deg ids with deg status
            big_rep['DEG ID Status'] = big_rep['deg_num'].map(deg_status)

            #delete deg ids without its status
            big_rep = big_rep[~big_rep['DEG ID Status'].isnull()]

            #get Assistant in the correct format
            big_rep['Assistant'] = big_rep['Assistant'].str.replace("  ", " ")
            big_rep['Assistant'] = big_rep['Assistant'].str.strip()
            big_rep['Assistant'] = np.where((big_rep['Assistant'].str.count("\(") > 1),
                                            big_rep['Assistant'].str[:-11],
                                            big_rep['Assistant'])
            rec = big_rep['Assistant'].str.split(',', expand=True)
            big_rep['Assistant'] = rec[1].str[1:-11]+" "+rec[0]+rec[1].str[-11:]

            #no app will be int
            big_rep['no_deg'] = pd.to_numeric(big_rep['no_deg'])

            # delete unnecessary columns
            big_rep = big_rep[['deg_num', 'no_deg', 'deg_id', 'DEG ID Status', 'ServiceTicket ID', 'Current Status', 'ServiceTicket Status', 'Assistant']]

            #STs which are assigned to one deg ID only (1, 2, 3, 4, 5)
            unique_ids = big_rep.drop_duplicates(subset=['deg_id'], keep=False)['ServiceTicket ID'].tolist()

            #create a new column for such cases
            big_rep['Unique_ID'] = np.where(big_rep['ServiceTicket ID'].isin(unique_ids), True, False)

            #crtete a column showing how many active STs each deg ID has
            big_rep['No. actives'] = big_rep.groupby('deg_id')['Current Status'].transform(count_statuses, report='active')

            #crtete a column showing how many completed STs each deg ID has
            big_rep['No. completed'] = big_rep.groupby('deg_id')['Current Status'].transform(count_statuses, report='completed')

            #crtete a column showing how many Canceled STs each deg ID has
            big_rep['No. canceled'] = big_rep.groupby('deg_id')['Current Status'].transform(count_statuses, report='canceled')

            big_rep = big_rep.sort_values(by='deg_num', ascending=False).reset_index(drop=True)

            #get list of deg IDs that are present in active_completed reports
            ids_with_sts = big_rep['deg_num'].drop_duplicates().tolist()

            #create a list of deg ids in transaction completed status which have only one serviceticket assigned and it is in completeds status
            degs_tr_completed_list = big_rep[(big_rep['DEG ID Status'] == 'Transaction Completed') & (big_rep['Unique_ID'] == True) & big_rep['No. completed'] == 1]['deg_num'].tolist()

            #from degs_fulls take deg ids that are in degs_tr_completed_list
            degs_tr_completed = degs_fulls[degs_fulls['DEG System ID'].isin(degs_tr_completed_list)].drop(['Status', 'Access Type'], axis=1).copy()

            degs_accepted = pd.concat([degs_accepted, degs_tr_completed], axis=0)

            return degs_accepted, degs_fulls, big_rep, actives_completed, ids_with_sts
        
    
        def manipulate(big_rep, degs_accepted, degs_fulls):
            """Fulls operations to get final results"""

            #create dictionaries from big_rep

            #lists required
            st_dicts = []
            status_dicts = []
            assistant_dicts = []
            correct_ids_dicts = []

            for x in range(1,6):
                #create a dictionary of ids assigned to one BR only
                unique_dict = big_rep[(big_rep['no_deg'] == x) & (big_rep['Unique_ID'] == True)].set_index('deg_num').to_dict()['ServiceTicket ID']
                #create a dictionary of ids which appear only once on active
                one_active_dict = big_rep[(big_rep['no_deg'] == x) &
                                        (big_rep['Unique_ID'] == False) &
                                        (big_rep['No. actives'] == 1) &
                                        (big_rep['Current Status'].isin(statuses_active))].set_index('deg_num').to_dict()['ServiceTicket ID']
                #create dictionary of ids which appear only once on completed
                one_completed_dict = big_rep[(big_rep['no_deg'] == x) &
                                        (big_rep['Unique_ID'] == False) &
                                        (big_rep['No. actives'] < 1) &
                                        (big_rep['No. completed'] == 1) &
                                        (big_rep['Current Status'].isin(statuses_completed))].set_index('deg_num').to_dict()['ServiceTicket ID']
                #create dictionary of ids which appear only once on canceled
                one_canceled_dict = big_rep[(big_rep['no_deg'] == x) &
                                        (big_rep['Unique_ID'] == False) &
                                        (big_rep['No. actives'] < 1) &
                                        (big_rep['No. completed'] < 1) &
                                        (big_rep['No. canceled'] == 1) &
                                        (big_rep['Current Status'].isin(statuses_canceled))].set_index('deg_num').to_dict()['ServiceTicket ID']
                #create dictionary of ids in Transaction Completed status which are unique and appear only once in completed
                one_completed_tr_completed_dict = big_rep[(big_rep['no_deg'] == x) &
                                        (big_rep['Unique_ID'] == True) &
                                        (big_rep['No. completed'] == 1) &
                                        (big_rep['DEG ID Status'] == 'Transaction Completed') &
                                        (big_rep['Current Status'].isin(statuses_completed))].set_index('deg_num').to_dict()['ServiceTicket ID']

                #merge all the above dictionaries
                merged_dict = {**unique_dict, **one_active_dict, **one_completed_dict, **one_canceled_dict, **one_completed_tr_completed_dict}

                stat_dict = big_rep[big_rep['no_deg'] == x].set_index('ServiceTicket ID')['ServiceTicket Status'].to_dict()
                assist_dict = big_rep[big_rep['no_deg'] == x].set_index('ServiceTicket ID')['Assistant'].to_dict()
                corr_dict = big_rep[big_rep['no_deg'] == x].set_index('ServiceTicket ID')['deg_num'].to_dict()
                
                st_dicts.append(merged_dict)

                status_dicts.append(stat_dict)
                assistant_dicts.append(assist_dict)
                correct_ids_dicts.append(corr_dict)

            #add new STs or update existing ones in degs in Accepted status -->sts_updated, new_degs_fulls
            #create a list for dfs
            list_of_dfs = []

            for (o, n, b) in (zip(old_sts, new_sts, st_dicts)):
                degs_rec = degs_accepted[['DEG System ID', o]].copy()
                degs_rec[n] = degs_rec['DEG System ID'].map(b)
                degs_rec = degs_rec[degs_rec[o] != degs_rec[n]].dropna(subset=[n]).drop(columns=o)
                list_of_dfs.append(degs_rec)
                sts_updated = ft.reduce(lambda left, right: pd.merge(left, right, on='DEG System ID', how='outer'), list_of_dfs)

            #identify deg ids with st id and its status but without assistant
            ids_no_assistant = []
            for(o, os, oa) in (zip(old_sts, old_statuses, old_assistants)):
                list_no_assistants = degs_fulls[(degs_fulls[o].notna()) &
                                                (degs_fulls[os].notna()) &
                                                (degs_fulls[oa].isnull())]['DEG System ID'].tolist()
                ids_no_assistant.append(list_no_assistants)
            ids_no_assistant = list(np.concatenate(ids_no_assistant))
            ids_no_assistant = list(set(ids_no_assistant))

            # create a copy of degs_fulls
            new_degs_fulls = degs_fulls.copy()

            #merge degs_fulls & sts_updated DFs (left)
            new_degs_fulls = new_degs_fulls.merge(sts_updated, how='left')

            for (o, n) in (zip(old_sts, new_sts)):
                #replace ST ID where necessary
                new_degs_fulls[o] = np.where((new_degs_fulls[o] != new_degs_fulls[n]) & (new_degs_fulls[n].notna()),
                                             new_degs_fulls[n],
                                             new_degs_fulls[o])
            
            #drop new columns added
            new_degs_fulls = new_degs_fulls.iloc[:,:18]

            #create a list of deg aproval IDs with duplicated servicetickets
            ids_with_duplicates = []
            for ob in old_sts:
                list_created = new_degs_fulls[(new_degs_fulls[ob].notna()) &
                                                (new_degs_fulls[ob].str.endswith('ST')) &
                                                (new_degs_fulls[ob].duplicated(keep=False)) &
                                                (new_degs_fulls[ob].str.len() == 7)]['DEG System ID'].tolist()
                ids_with_duplicates.append(list_created)
            ids_with_duplicates = list(np.concatenate(ids_with_duplicates))
            ids_with_duplicates = list(set(ids_with_duplicates))

            #Clear STs assigned to different deg ids or not assigned to any id
            new_degs_corrected = new_degs_fulls.copy()

            for (o, c) in (zip(old_sts, correct_ids_dicts)):
                #replace ST ID where necessary (maybe additional condition (isin actives & completed originals??))
                new_degs_corrected['deg_assigned'] = new_degs_corrected[o].map(c)
                new_degs_corrected[o] = np.where((new_degs_corrected['DEG System ID'] != new_degs_corrected['deg_assigned']) &
                                                    (new_degs_corrected[o].notna()),
                                                    np.nan,
                                                    new_degs_fulls[o])
                
            new_degs_corrected = new_degs_corrected.iloc[:,:-1]

            #Update statuses and Assistants in new_degs_corrected
            full_degs = new_degs_corrected.copy()

            for(o, os, orr, sd, rd) in (zip(old_sts, old_statuses, old_assistants, status_dicts, assistant_dicts)):
                full_degs[os] = full_degs[o].map(sd)
                full_degs[orr] = full_degs[o].map(rd)

            full_degs = full_degs[all_columns]
            return full_degs, ids_no_assistant, ids_with_duplicates
            

        def final_reports(full_degs, degs_fulls, actives_completed, big_rep, ids_with_duplicates, ids_with_sts, ids_no_assistant):
            """Prepare final reports"""

            #df showing all differences between degs_fulls & full_degs
            final_all = pd.merge(full_degs, degs_fulls, indicator=True, how='outer').query('_merge=="left_only"').drop('_merge', axis=1).iloc[:,:-2]

            # data with differences minus ids with no recruiter added
            final_blanks = final_all[~final_all['DEG System ID'].isin(ids_no_assistant)]

            #list of columns from ServiceTicket ID 1 to Assistant 5
            columns_required = all_columns_plus[3:]

            # create df from final blanks with deg IDs which have duplicated service tickets
            duplicated_blanks = final_blanks[final_blanks['DEG System ID'].isin(ids_with_duplicates)]                              
            duplicated_blanks = duplicated_blanks.loc[~duplicated_blanks.index.isin(duplicated_blanks.dropna(subset=columns_required,
                                                                                                             how='all').index)]

            #remove IDs that have blanks only
            final_no_blanks = final_blanks.dropna(subset=columns_required, how='all')

            #concatenate final_no_blanks & duplicated_blanks to get a final result
            final_upload = pd.concat([final_no_blanks, duplicated_blanks], ignore_index=True)
            #remove deg IDs wthat should be excluded
            final_upload = final_upload[~final_upload['DEG System ID'].isin(degs_excluded)]

            now = datetime.now()
            date = now.strftime('%d-%b-%Y_%H_%M')

            #prepare full file for analysis
            with pd.ExcelWriter('results/FULL deg ST Status '+date+'.xlsx') as writer:
                degs_fulls.to_excel(writer, sheet_name='Degs_Original_Limited', index=False) #degs that can be potentially updated
                actives_completed.to_excel(writer, sheet_name='Act_Cl_Original', index=False) #original active & completed merged
                big_rep.to_excel(writer, sheet_name='Act_Cl_Modified', index=False) #combination of actives and completed limited to deg IDs required
                final_all.to_excel(writer, sheet_name='All_Differences', index=False) #all deg ids that were updated
                final_blanks.to_excel(writer, sheet_name='Final_with_Blanks', index=False) #list of ids to upload containing also blanks
                final_upload.to_excel(writer, sheet_name='Final_Upload', index=False) #list of ids to upload

            len_full = len(final_upload)
            div = len(final_upload) / row_limit
            mult = int(div)
            sub = len_full - (mult * row_limit)

            #if number of rows is more than row_limit, split the file into parts
            if len_full <= row_limit:
                final_upload.to_excel("results/deg st status "+ date+" ({:02}).xlsx".format(1), sheet_name='DegEntry', index=False)
            elif sub == 0:
                i=1
                for huge_df in np.array_split(final_upload, len_full // row_limit):
                    huge_df.to_excel("results/deg st status " +date+" ({:02}).xlsx".format(i), sheet_name='DegEntry', index=False)
                    i += 1
            else:
                new_final_upload = final_upload[:-sub]
                rest = final_upload[-sub:]
                i=1
                for huge_df in np.array_split(new_final_upload, len(new_final_upload) // row_limit):
                    huge_df.to_excel("results/deg ST Status " +date+" ({:02}).xlsx".format(i), sheet_name='DegEntry', index=False)
                    i += 1
                rest.to_excel("results/deg st status " +date+" ({:02}).xlsx".format(i), sheet_name='DegEntry', index=False)
            
            print("\nNumber of records to upload:", len_full)

        panel = "control panel.xlsx"
        try:
            mapping = pd.read_excel(panel, sheet_name="mapping")
        except FileNotFoundError:
            print(panel, "not found")
            message = "\n\n\n\nControl Panel not found!\n"
            self.text_update(message)
            msb.showerror(title="Error", message="control panel.xlsx not found!")
            return None
        
        #list of ST Status names excluding website visit
        mapp = [x for x in mapping['ST Status Name'] if str(x) != 'website visit']
        #dictionary: key==ST Status name, value==category
        status_dict = dict(zip(mapping['ST Status Name'], mapping['Category']))
        #take limit of rows in final file(s)
        row_limit = int(mapping['Row Limit'][0])
        #min and max values for deg IDs to be considered
        min_deg = int(mapping['Min DEG ID'][0])
        max_deg = int(mapping['Max DEG ID'][0])

        #deg ids that should be excluded from teh final uplaod
        degs_excluded = [x for x in mapping['Excluded DEG IDs'].tolist() if str(x) != 'nan']

        #columns required in the final file
        all_columns = ['DEG System ID', 'ServiceTicket ID 1', 'ServiceTicket Status 1', 'Assistant 1',
                        'ServiceTicket ID 2', 'ServiceTicket Status 2', 'Assistant 2',
                        'ServiceTicket ID 3', 'ServiceTicket Status 3', 'Assistant 3',
                        'ServiceTicket ID 4', 'ServiceTicket Status 4', 'Assistant 4',
                        'ServiceTicket ID 5', 'ServiceTicket Status 5', 'Assistant 5']

        #all columns + status & Access Type
        all_columns_plus = ['DEG System ID', 'Status', 'Access Type',
                        'ServiceTicket ID 1', 'ServiceTicket Status 1', 'Assistant 1',
                        'ServiceTicket ID 2', 'ServiceTicket Status 2', 'Assistant 2',
                        'ServiceTicket ID 3', 'ServiceTicket Status 3', 'Assistant 3',
                        'ServiceTicket ID 4', 'ServiceTicket Status 4', 'Assistant 4',
                        'ServiceTicket ID 5', 'ServiceTicket Status 5', 'Assistant 5']

        #list of columns required to process data
        old_sts = ['ServiceTicket ID 1', 'ServiceTicket ID 2','ServiceTicket ID 3', 'ServiceTicket ID 4', 'ServiceTicket ID 5']
        new_sts = ['New_ST_ID 1', 'New_ST_ID 2', 'New_ST_ID 3', 'New_ST_ID 4', 'New_ST_ID 5']
        old_statuses = ['ServiceTicket Status 1', 'ServiceTicket Status 2','ServiceTicket Status 3', 'ServiceTicket Status 4', 'ServiceTicket Status 5']
        new_statuses = ['New_ST_Status 1', 'New_ST_Status 2', 'New_ST_Status 3', 'New_ST_Status 4', 'New_ST_Status 5']
        old_assistants = ['Assistant 1', 'Assistant 2', 'Assistant 3', 'Assistant 4', 'Assistant 5']
        new_assistants = ['New_Assistant 1', 'New_Assistant 2', 'New_Assistant 3', 'New_Assistant 4', 'New_Assistant 5']

        #create lists for active, completed and canceled statuses
        statuses_active = ['Active', 'Approved', 'Suspended', 'Pending']
        statuses_canceled = ['Canceled', 'Declined', 'Rejected', 'Deleted']
        statuses_completed = ['Completed']

        message = "\n\n\n\nPlease wait...\n"
        self.text_update(message)
        file1 = askopenfile(title = "Daily serviceticket rep. active", mode = "r",
                             filetypes = [("CSV Files", "*.csv")])
        file2 = askopenfile(title = "Daily serviceticket rep. completed CANCELED", mode = "r",
                             filetypes = [("CSV Files", "*.csv")])
        file3 = askopenfile(title = "DEG System Report", mode = "r",
                             filetypes = [("Excel Files", "*.xlsx")])

        start_time = datetime.now()

        if file1 and file2 and file2 is not None:
            try:
                actives = pd.read_csv(file1.name, dtype = {'DEG System ID' : str})
                completed = pd.read_csv(file2.name, dtype = {'DEG System ID' : str})
                degs = pd.read_excel(file3.name, sheet_name="DegEntry")
            except pd.errors.ParserError:
                print("Incorrect format")
                message = "\n\n\nThe files must have CSV extension.\n"
                message += "\nResave the files and try again.\n"
                self.text_update(message)
                msb.showerror(title="Error", message="Make sure the fikes are saved as CSV!")
                return None
            file1.close()
            file2.close()
            file3.close()

            #dictionary: key:deg_id, value:deg status
            deg_status = dict(zip(degs['DEG System ID'], degs['Status']))

            degs_accepted, degs_fulls, big_rep, actives_completed, ids_with_sts = prepare_reports(degs, actives, completed, deg_status)
            full_degs, ids_no_assistant, ids_with_duplicates = manipulate(big_rep, degs_accepted, degs_fulls)
            final_reports(full_degs, degs_fulls, actives_completed, big_rep, ids_with_duplicates, ids_with_sts, ids_no_assistant)

            message = "\n\n\n\nJob complete!\n"
            self.text_update(message)
            end_time = datetime.now()
            print('\nJob complete. Duration: {}'.format(end_time - start_time))
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
root.title("DEG - ST Status Update")
root.geometry("554x172")
root.resizable(False, False)
app = Application(root)

root.rowconfigure(0, minsize=75, weight=1)
root.columnconfigure(1, minsize=50, weight=3)

root.mainloop()