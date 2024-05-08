import os
import sys

script_path = os.path.realpath(__file__)
parent_path = os.path.dirname(script_path)
main_folder, subfolder = os.path.split(parent_path)
src_folder = os.path.join(os.path.sep,main_folder,"Src")
sys.path.insert(0, src_folder)

from SAP_Automation_Class import SAP_Automation

class CJI3_Automation_Class(SAP_Automation):

    def process_transaction(self):
        import datetime
        today = datetime.date.today()
        first_day = today.replace(day=1)
        end = first_day - datetime.timedelta(days=1)
        start = (first_day - datetime.timedelta(days=30)).replace(day=1).replace(month=1)

        start_str=start.strftime("%d.%m.%Y")
        end_str=end.strftime("%d.%m.%Y")

        print("Taking the range from {} until {} .".format(start_str,end_str))

        session=self.PRD_session

        session.SendCommand("/nCJI3")

        # Select Transaction paramaterss 

        #Select Database Profile
        session.findById("wnd[1]").sendVKey(4)
        session.findById("wnd[2]/usr/lbl[14,3]").setFocus()
        session.findById("wnd[2]/usr/lbl[14,3]").caretPosition = 8
        session.findById("wnd[2]").sendVKey(2)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Choose Range
        session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").caretPosition = 0
        session.findById("wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL").select()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-ILOW_I[1,0]").text = "56000000" # Range lower bound
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").text = "56999999" # Range upperbound
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").caretPosition = 8
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # Choose Date
        session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").text = start_str # begin date
        session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").text = end_str # end date
        session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").setFocus()

        # Set Rapport Lay-out
        session.findById("wnd[0]/usr/ctxtP_DISVAR").setFocus()
        session.findById("wnd[0]/usr/ctxtP_DISVAR").caretPosition = 4
        session.findById("wnd[0]").sendVKey(2)
        session.findById("wnd[0]").sendVKey(2)
        session.findById("wnd[0]").sendVKey(4)
        session.findById("wnd[1]").close()
        session.findById("wnd[0]/usr/ctxtP_DISVAR").text = "/MLC/IR"
        session.findById("wnd[0]/usr/ctxtP_DISVAR").caretPosition = 7
        session.findById("wnd[0]").sendVKey(0)

        # Set amount of hits
        session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").caretPosition = 2
        session.findById("wnd[0]/usr/btnBUT1").press()
        session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "20000" #Amount of hits
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        return
    
    def export_to_excel(self):
        session=self.PRD_session

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\{}\Downloads".format(self.user)
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "CJI3.XLSX" #set filename
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        return 
    

#%%  Actual Code

print("Initialising code . . .")
CJI3=CJI3_Automation_Class()

print("Opening/Connecting with SAP . . .")
CJI3.connect_sap()

print("Processing . . .")
CJI3.process_transaction() # Dynamic Dispatch method aka Overide method.

print("Starting export to Excel . . .")
CJI3.export_to_excel()  # Dynamic Dispatch method aka Overide method.

print("Done with Transaction, closing SAP . . .")
CJI3._close_sap()

#%% Data mangeling

print("Starting data processing . . .")
import pandas as pd
from time import sleep
import subprocess
import numpy as np
import warnings

warnings.filterwarnings("ignore")

user = CJI3.user
path="C:/Users/{}/Downloads/CJI3.xlsx".format(user)

sleep(4)
subprocess.call("taskkill /f /im EXCEL.EXE", shell=True)


data=pd.read_excel(path)

data['month'] = data['Posting Date'].dt.month

# Create and configure list of unique object numbers
Object_list=pd.unique(data["Object"])
Object_list=pd.DataFrame(Object_list,columns=["Object"])
Object_list.dropna(inplace=True)
Object_list["Object"]=Object_list["Object"].astype(int)
Object_list.set_index(["Object"],inplace=True)

# Create and configure list of unique months
unique_months=pd.unique(data["month"])
unique_months=unique_months[~np.isnan(unique_months)].astype(int)
unique_months.sort()

for month in unique_months:
    Object_list[month]=np.nan

# For every month, get monthly data, sum all the currency per Object number and append to Object list
for month in unique_months:
    data_filterd=data[data["month"]==month]

    #grouped=data_filterd.groupby(by=["Object"]).sum()

    for object in pd.unique(data_filterd["Object"]):
        data_object=data_filterd[data_filterd["Object"]==object]
        sum=np.sum(data_object["Object"])
        Object_list[object,month]=sum

    #grouped[month]=grouped["Val.in rep.cur."]
    
    #grouped.index = grouped.index.astype(int)
    #Object_list=pd.concat([Object_list,grouped[month]],axis=1)
    

Output=Object_list.fillna("-")
print(Output)
Output.to_excel("C:/Users/klabbf/Downloads/WOMS_Actually.xlsx")

print("Export of file completed, programm shutting down.")