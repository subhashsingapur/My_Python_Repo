from jira import JIRA
import datetime
import pandas as pd
import os
from selenium import webdriver
import time
import logging
import getpass
import sys

class jira_creation():
    def __init__(self):
        self.username = getpass.getuser()
        self.pswd = input("please provide the system password for jira auth :")
        self.jira = JIRA("https://jira-adas.zone2.agileci.conti.de",
                         auth=(self.username, self.pswd))  # a username/password tuple
        self.summary =None
        self.assignee = None
        self.start_date = None
        self.end_date = None
        self.description = None
        self.story_points = None
        self.priority = None
        self.status = None
        self.issue_dict = {}
        self.ticket_list = []
        self.write = False
        print("Connecting to Jira Server")

    def create_jira(self, summary, assignee, start_date, Due_date, description, story_points, priority, status,Fix_ver,type_of):
        if not len(str(Fix_ver)):
            raise Exception("Label/Fix version Field is empty in planner,Please Add and export excel again")
        self.issue_dict = {
            'project': 'ARS540BW',
            'summary': summary,
            'description': description,
            'issuetype': {'name': 'Story'},
            # 'reporter':'uidk3496',
            "priority": {"id": str(priority)},
            "labels": [type_of],
            "assignee": {'name': assignee},
            "customfield_10105": start_date,
            # "customfield_10106": end_date,
            "customfield_10006": story_points,
            "fixVersions":[{'name':Fix_ver}],
            "duedate": Due_date
        }
        self.issue = self.jira.create_issue(fields=self.issue_dict)
        self.jira.transition_issue(self.issue, transition=status)
        self.ticket_id = self.issue.key
        print(" Newly Created : Ticket ID : {}, assignee: {},  status : {}".format(self.ticket_id,self.issue.fields.assignee,status))
        self.jira.add_watcher(self.issue, assignee)
        return self.ticket_id

    def date(self, start_date):
        self.date_input = start_date.replace('/', '-')
        self.datetimeobject = datetime.datetime.strptime(self.date_input, '%m-%d-%Y')
        self.new_start_date = self.datetimeobject.strftime('%Y-%m-%d')
        return self.new_start_date

    def get_prio(self, priority):
        self.prio_dict = {
            'Urgent' : 2,
            'Important': 2,
            'Medium': 3,
            'low': 4
        }
        return self.prio_dict[priority]

    def get_userid(self, assignee):
        self.assignee = str(assignee).split(';')
        # print(self.assignee)
        if len(self.assignee)>1:
            self.assignee = str((self.assignee)[0]).split("(")
            self.user_id = self.assignee[1].replace(')', '')
            return self.user_id.strip('\]')
        else:
            self.assignee = str(self.assignee).split("(")
            self.user_id = self.assignee[1].replace(')', '')
            return self.user_id.replace('\']','')

    def get_status(self, status):
        self.status_dict = {
            "Not started": "Planned",
            "In progress": "In Progress",
            "Completed": "Done",
        }
        return self.status_dict[status]

    def overwrite_excel(self, my_excel,my_df,my_status,ticket_list):
        self.my_dict = {
                        'Task ID': my_df,
                        'Progress':my_status,
                        'Ticket ID':ticket_list
                        }
        self.df1=pd.DataFrame(self.my_dict)
        self.df1.to_excel(my_excel,index=False)

    def update_jira(self,ticket_id,status,summary,fix_version,type_of,end_date):
        self.myissue = self.jira.issue(ticket_id)
        self.myissue.update(fields={'description':summary,"customfield_10105": end_date,"labels": [type_of],"fixVersions":[{'name':fix_version}]})
        if status == "Cancelled" or status == "On Hold":
            self.jira.transition_issue(self.myissue, transition=str(status),customfield_10529=input("Please Provide the reason for cancelling"))

        elif str(self.myissue.fields.status) == str(status):
            print("Already finished ",ticket_id)
        else:
            try:
                self.jira.transition_issue(self.myissue, transition=str(status))
                print("Updated :  Ticket ID = {} assignee = {}  Updated status = {}".format(ticket_id,
                                                                                        self.myissue.fields.assignee,
                                                                                        status))
            except:
                pass
    def compare_status(self,planner_Excel,mem_excel,type_of):
        self.updated_status = []
        self.updated_Ticket_id =[]
        self.update_taskid=[]
        df = planner_Excel
        df_mem = pd.read_excel(mem_excel)
        for ind in df.index:
            for ind_1 in df_mem.index:
                self.same_taskid = False
                if df['Task ID'][ind] == df_mem['Task ID'][ind_1]:
                    if df['Progress'][ind] == df_mem['Progress'][ind_1]:
                        if str(df['Progress'][ind]) != str("Completed"):
                            self.same_taskid=True
                            self.updated_status.append(df['Progress'][ind])
                            self.updated_Ticket_id.append(df_mem['Ticket ID'][ind_1])
                            self.update_taskid.append(df['Task ID'][ind])
                            print("No Change : Ticket id {} is in same state : {} for assignee {}".format(df_mem['Ticket ID'][ind_1],
                                                                                            df['Progress'][ind],df['Assigned To'][ind]))
                            break

                    else:
                        self.update_jira(df_mem['Ticket ID'][ind_1],self.get_status(df['Progress'][ind]),df["Description"][ind],df["Labels"][ind],type_of,self.date(df['Created Date'][ind]))
                        self.updated_status.append(df['Progress'][ind])
                        self.updated_Ticket_id.append(df_mem['Ticket ID'][ind_1])
                        self.update_taskid.append(df['Task ID'][ind])

            # if ind_1!=0:
            if not df['Progress'][ind] == "Completed" and not self.same_taskid:
                self.updated_Ticket_id.append(
                    self.create_jira(df['Task Name'][ind], self.get_userid(df['Assigned To'][ind]),
                                     self.date(df['Created Date'][ind]), self.date(df['Due Date'][ind]),
                                     df['Description'][ind], 1, self.get_prio(df['Priority'][ind]),
                                     self.get_status(df['Progress'][ind]), df["Labels"][ind],type_of))
                self.update_taskid.append(df['Task ID'][ind])
                self.updated_status.append(df['Progress'][ind])

        self.overwrite_excel(mem_excel,self.update_taskid,self.updated_status,self.updated_Ticket_id)

    def cleanup_Excel(self,input_excel,type_of):
        self.input_excel = input_excel
        self.xls = pd.ExcelFile(self.input_excel)
        self.df_ = self.xls.parse(skiprows=4, index_col=None)
        # self.df_.to_excel(self.input_excel,index=False)
        some_obj.check_all_fields(self.df_)
        self.read_Excel(self.df_,type_of)

    def create_excel(self,memory_Excel):
        with open(memory_Excel,'w') as fp:
            pass

    def read_Excel(self, input_df,type_of):
        self.first_time = True
        self.input_df = input_df
        self.mem_excel_path=r"\\lifs010\data\BCPSENCAP\_Validation\EBA\99_temp\Subhash-Temp\Memory_Excel\new_mem_excel.xlsx"
        self.taskId =[]
        self.status_copy =[]
        print("Ticket Statistics")
        for ind in self.input_df.index:
            if not os.path.exists(self.mem_excel_path):
                if not self.input_df['Progress'][ind] == "Completed":
                    self.summary = self.input_df['Task Name'][ind]
                    self.assignee = self.get_userid(self.input_df['Assigned To'][ind])  # single assignee is needed
                    self.start_date = self.date(self.input_df['Created Date'][ind])
                    self.Due_date = self.date(self.input_df['Due Date'][ind])
                    self.description = self.input_df['Description'][ind]
                    self.story_points = 1
                    self.priority = self.get_prio(self.input_df['Priority'][ind])
                    self.status = self.get_status(self.input_df['Progress'][ind])
                    self.fix_ver = self.input_df["Labels"][ind]
                    self.taskId.append(self.input_df['Task ID'][ind])
                    self.status_copy.append(self.status)
                    self.write = True
                    self.ticket_list.append(self.create_jira(self.summary, self.assignee, self.start_date, self.Due_date, self.description,
                            self.story_points, self.priority, self.status,self.fix_ver,type_of))
            else:
                self.compare_status(self.input_df,self.mem_excel_path,type_of)
                os.remove(self.input_excel)
                break
        if self.write:
            self.create_excel(self.mem_excel_path)
            self.overwrite_excel(self.mem_excel_path,self.taskId, self.status_copy,self.ticket_list)
            os.remove(self.input_excel)

class Planner2Excel():
    def download_excel(self):
        self.edge_path = r"D:\BPL_SW\planner2jira\edgedriver_win64\msedgedriver.exe"
        self.driver = webdriver.Edge(self.edge_path)
        self.driver.get('https://tasks.office.com/continental.onmicrosoft.com/Home/PlanViews/BJY1F-d300S8Cwpib_NhGpcAFrES?Type=PlanLink&Channel=Link&CreatedTime=637525354108450000')
        time.sleep(20)
        print("WebPage is Loaded")
        self.excel_link = self.driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[4]/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div[2]/div/span/button')
        self.excel_link.click()
        self.excel_link = self.driver.find_element_by_xpath('/html/body/div[2]/div/div/div/div/div/div/ul/li[9]/button')
        self.excel_link.click()
        print("Planner Tasks will be exported to Excel in few secs")
        # self.driver.quit()

    def check_all_fields(self,df):
        to_exit=False
        for idx in df.index:
            if str(df['Progress'][idx]) != "Completed":
                if str(df['Assigned To'][idx]) == "nan":
                    logging.error("Please assign one assignee for task name {} and then retry".format(df['Task Name'][idx]))
                    to_exit=True
                elif str(df['Due Date'][idx]) == "nan":
                    logging.error("Please assign the Due Date for task name {} and then retry".format(df['Task Name'][idx]))
                    to_exit = True
                elif str(df['Labels'][idx]) == "nan":
                    logging.error("Please give the Fix version for task name {} and then retry".format(df['Task Name'][idx]))
                    to_exit = True
                else:
                    pass
        if to_exit:
            if input("Press any Key to continue.."):
                exit(0)


if '__main__' == __name__:
    # Planner2Excel().download_excel()
    # time.sleep(20)
    input_excel_path = input('Please input the task planner excel path :')
    input_excel_path=input_excel_path.strip('"')
    # for sometime in range(10):
    #     if not os.path.exists(input_excel_path):
    #         wait_time=20
    #         time.sleep(wait_time)
    # print("Excel is exported Successfully.")
    if os.path.exists(input_excel_path):
        some_obj = Planner2Excel()
        jira_obj = jira_creation()
        if "UC" in os.path.split(input_excel_path)[1]:
            jira_obj.cleanup_Excel(input_excel_path,"EBA_Usecase")
        else:
            jira_obj.cleanup_Excel(input_excel_path, "EBA_Endurance")
    else:
        logging.error("No input excel")
    if input("Press any Key to continue.."):
        sys.exit(0)