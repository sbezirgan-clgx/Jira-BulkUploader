import difflib
from jira import JIRA
import requests
import json
import pandas as pd
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
    if issue_status.lower() == 'fixesdone' or issue_status.lower() == 'fixes done':
        return 'Fixed'
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
def startConnection():
    '''Jira Server Connection'''
    jiraOptions = {'server': "https://jira-corelogic.valiantys.net"}
    jira = JIRA(options=jiraOptions, basic_auth=("Internal-EDG-SA-BulkDefects", "iXyM*8W!s84&"))
    return jira


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    jira= startConnection()
    my_list = read_excel_file("Example.xlsx", [0,2,3,4])
    for issue_info in my_list:
        issue_id = issue_info[0]
        issue = jira.issue(issue_id)
        issue_status = issue_info[1]
        issue_status_new = get_most_similar_issue_status_from_transition_name_list(jira,issue,issue_status)
        issue_comment_tcs = issue_info[2]
        issue_comment_onshore = issue_info[3]
        bool_tcs_comment = comment_cross_check_excel(jira, issue_id, issue_comment_tcs)
        bool_onshore_comment = comment_cross_check_excel(jira, issue_id, issue_comment_onshore)
        if issue_status_new is not None:
            set_issue_status_by_transition_name(jira,issue,issue_status_new)
        if issue_comment_tcs is not None and bool_tcs_comment is False:
            add_comment_to_an_issue(jira, issue_id, "(TCS) "+issue_comment_tcs)
        if issue_comment_onshore is not None and bool_onshore_comment is False:
            add_comment_to_an_issue(jira, issue_id, "(Onshore) "+issue_comment_onshore)

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





