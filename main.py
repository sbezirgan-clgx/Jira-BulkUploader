import difflib
import os
from jira import JIRA
import requests
import json
import pandas as pd
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import os
import openpyxl
BACKGROUND_COLOR = "#B1DDC6"
JIRA_USERNAME = os.getenv("JIRA_USERNAME")
JIRA_PASSWORD = os.getenv("JIRA_PASSWORD")
records_count = 1
window = Tk()

#Progress Bar
prg1 = ttk.Progressbar(window,orient = HORIZONTAL,
        value=0,length = 300, mode = 'determinate')
prg1.grid(row=7, column=1, columnspan=2)
progress_label = Label(text="Progress")
progress_label.grid(row=6, column=1, columnspan=2)


def get_transition_id_by_name(jira: JIRA, issue, name: str):
    transitions = jira.transitions(issue)
    transition_list = [(t['id'], t['name']) for t in transitions]
    print(transition_list)
    transition_list_iterator = filter(lambda x: (x[1] == name), transition_list)
    filtered_transition_list = list(transition_list_iterator)
    print(filtered_transition_list)
    return filtered_transition_list[0][0]


def get_transition_name_list(jira:JIRA, issue):
    transitions = jira.transitions(issue)
    transition_list = [t['name'] for t in transitions]
    return transition_list


def get_most_similar_issue_status_from_transition_name_list(jira:JIRA,issue,issue_status:str):
    if issue_status.lower() == 'fixesdone' or issue_status.lower() == 'fixes done' or issue_status.lower() == 'correction done' or issue_status.lower() == 'already corrected':
        return 'Fixed'
    elif issue_status.lower() == 'transaction not found' or issue_status.lower() == 'field not found':
        return 'Not in Scope'
    elif issue_status.lower() == 'transaction deleted':
        return 'Non-Issue'
    transition_name_list = get_transition_name_list(jira,issue)
    most_similar_issue_status_list = difflib.get_close_matches(issue_status, transition_name_list)
    if len(most_similar_issue_status_list) > 0:
        return most_similar_issue_status_list[0]
    return 'Non-Issue'


def set_issue_status_by_transition_name(jira:JIRA, issue, transition_name:str):
    transition_id=get_transition_id_by_name(jira,issue,transition_name)
    jira.transition_issue(issue, transition_id)


def read_excel_file(file_name:str, require_cols:list):
    required_df= pd.read_excel(file_name, usecols=require_cols)
    return list(required_df.itertuples(index=False, name=None))


def add_comment_to_an_issue(jira:JIRA, issueID, comment:str):
    myissue= jira.issue(issueID)
    jira.add_comment(myissue, comment)


def get_comment_list_from_an_issue(jira:JIRA, issueID):
    '''Returns a list of comment bodies for the Jira ticket'''
    issue = jira.issue(issueID)
    comments1 = issue.fields.comment.comments
    comment_list = []
    for comment in comments1:
        comment_list.append(comment.body)
    return comment_list


def comment_cross_check_excel(jira:JIRA, issueID,excel_comment:str):
    '''Returns true if the comment already exists in Jira'''
    comment_list= get_comment_list_from_an_issue(jira,issueID)
    for comment in comment_list:
        if excel_comment == comment:
            return True
            break
    return False


def start_connection():
    '''Jira Server Connection'''
    jiraOptions = {'server': "https://jira-corelogic.valiantys.net"}
    jira = JIRA(options=jiraOptions, basic_auth=(JIRA_USERNAME, str(password_entry.get())))
    return jira


def update_screen(label:Label, count:int, length:list):
    label.config(text=f"{str(count)}/{str(len(length))}")
    window.after(500, update_screen)


# Press the green button in the gutter to run the script.
def upload_records():
    if len(isc_entry.get()) == 0 or len(file_entry.get()) == 0 or len(username_entry.get()) == 0:
        messagebox.showinfo(title="Fill All The Fields", message="Please make sure you haven't left any fields empty.")
    if password_entry.get() != JIRA_PASSWORD and len(password_entry.get()) != 0:
        messagebox.showinfo(title="Password Error", message="Incorrect Password")
    jira = start_connection()
    global records_count
    records_count = 1
    try:
        my_list = read_excel_file(rf"C:\Users\{str(isc_entry.get())}\OneDrive - CoreLogic Solutions, LLC\Desktop\{str(file_entry.get())}.xlsx", [0,1,2])
    except PermissionError:
        messagebox.showinfo(title="ERROR", message="Please save and exit the Excel File before proceeding")
    else:
        for issue_info in my_list:
            issue_id = issue_info[2]
            issue = jira.issue(issue_id)
            issue_status = issue_info[0]
            issue_status_new = get_most_similar_issue_status_from_transition_name_list(jira,issue,issue_status)
            issue_comment_tcs = issue_info[1]
            #issue_comment_onshore = issue_info[3]
            bool_tcs_comment = comment_cross_check_excel(jira, issue_id, excel_comment=f"(TCS) {issue_comment_tcs}")
            #bool_onshore_comment = comment_cross_check_excel(jira, issue_id, excel_comment=f"(Onshore) {issue_comment_onshore}")
            print(bool_tcs_comment)
            #print(bool_onshore_comment)
            if issue_status_new is not None:
                set_issue_status_by_transition_name(jira,issue,issue_status_new)
            if issue_status_new == 'Not in Scope':
                if issue_status.lower() == 'field not found':
                    issue_comment_tcs = 'Field not found in THOR.'
                    bool_tcs_comment = False
                elif issue_status.lower() == 'transaction not found':
                    issue_comment_tcs = 'Transaction not found in THOR.'
                    bool_tcs_comment =False
            if issue_status_new == 'Fixed':
                issue_comment_tcs = 'Issue Fixed in THOR'
                bool_tcs_comment = False
            if issue_status_new == 'Non-Issue':
                issue_comment_tcs = 'Transaction deleted in THOR.'
                bool_tcs_comment = False
            if issue_comment_tcs is not None and bool_tcs_comment is False:
                add_comment_to_an_issue(jira, issue_id, "(TCS) "+issue_comment_tcs)
            #if issue_comment_onshore is not None and bool_onshore_comment is False:
                #add_comment_to_an_issue(jira, issue_id, "(Onshore) "+issue_comment_onshore)
            progress_label.config(text=f"{records_count}/{len(my_list)}")
            prg1.config(value=(100*records_count)/len(my_list))
            records_count += 1
            window.update_idletasks()
            if records_count == len(my_list)+1:
                messagebox.showinfo(title="Success", message="All the files have been successfully uploaded")



#-----------------------GUI---------------------#


window.title("JIRA-Uploader")
window.config(padx=50, pady=50)
canvas = Canvas(width=200, height=200)
logo_img = PhotoImage(file="Safeimagekit-resized-img.png")
canvas.create_image(100, 100, image=logo_img)
canvas.grid(row=0, column=1)


#Labels
isc_label = Label(text="ISC:")
isc_label.grid(row=1, column=0)
file_label = Label(text="File Name:")
file_label.grid(row=2, column=0)
username_label = Label(text="Username for Jira:")
username_label.grid(row=3, column=0)
password_label = Label(text="Password for Jira:")
password_label.grid(row=4, column=0)

#Entries
isc_entry = Entry(width=25)
isc_entry.grid(row=1, column=1)
isc_entry.focus()
file_entry = Entry(width=25)
file_entry.grid(row=2, column=1)
username_entry = Entry(width=35)
username_entry.grid(row=3, column=1, columnspan=2)
username_entry.insert(0, "Internal-EDG-SA-BulkDefects")
username_entry.config(state="disabled")
password_entry = Entry(width=21,show="*")
password_entry.grid(row=4, column=1)

# Buttons
add_button = Button(text="Update Status and Comment", width=36, command=upload_records,bg="blue", fg="white")
add_button.grid(row=5, column=1, columnspan=2)
add_button.config(padx=5, pady=5)


#-----------------------GUI---------------------
window.mainloop()








    # add_comment_to_an_issue(jira,'ETQA-5897','My test Comment')
    # set_issue_status_by_transition_name(jira,issue,'Closed')
    #print("/////////////////")

    #print(my_list)
    #issue.update(description= "Change Description")
    #transitions = jira.transitions(issue)
    #transition_list = [(t['id'], t['name']) for t in transitions]
    #print(transition_list)
    #target_status ="Closed"
    #transition_list_iterator = filter(lambda x: (x[1] == "Closed"), transition_list)
    #filtered_transition_list = list(transition_list_iterator)
    #print(filtered_transition_list)
    #transition_id= get_transition_id_by_name(jira,issue,'Closed')
    #print(transition_id)
    #jira.transition_issue(issue, '31')
    #print(issue)

    #summary = issue.fields.summary

    #print(summary)
    #print(issue.fields)

    #credentials ="Bearer Mzg4NjIxMDUzNDI2OoONxH4ZzfGiN6ndnGOD9/2rRMXi"
#Trying to get all issues for one project
    #headers={
        #"Accept": "application/json",
        #"Content-Type": "application/json",
        #"Authorization": credentials
    #}
    #projectKey ="ETQA"
    #url ="https://jira-corelogic.valiantys.net/rest/api/2/search?jql=project=" + projectKey + "&maxResults=1"
    #response=requests.request("GET",url,headers=headers)
    #print(response.json())

    #require_cols = [0,1,2,3]

    # only read specific columns from an excel file
    #required_df = pd.read_excel('Example.xlsx', usecols=require_cols)
    #print(type(required_df))
    #print(required_df)
    #Reading column names
    #for col in required_df.columns:
        #print (col)
    #Puts the columns names to a list
    #print(list(required_df.columns.values.tolist()))
    #For given column getting the cell value
    #print(required_df['Key'].loc[required_df.index[0]])
    #print(required_df['State'].values[0])
    #print(len(required_df))





