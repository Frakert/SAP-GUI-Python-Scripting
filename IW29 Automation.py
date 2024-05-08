from SAP_Automation_Class import SAP_Automation
import pandas as pd
from time import sleep

class IW29_Automation_Class(SAP_Automation):

    def open_transaction(self):
        self._PRD_session.SendCommand("/nIW29")
        return
        
    def process_transaction(self):
       session=self._PRD_session

       # Select Transaction paramaterss 
       session.findById("wnd[0]").resizeWorkingPane(210,28,"false")
       session.findById("wnd[0]/usr/chkDY_MAB").selected = "true" # include completed transactions
       session.findById("wnd[0]/usr/ctxtDATUV").text = "1.1.2023" # begin date
       session.findById("wnd[0]/usr/ctxtDATUV").setFocus()
       session.findById("wnd[0]/usr/ctxtDATUV").caretPosition = 3
       session.findById("wnd[0]/tbar[1]/btn[8]").press()
       return
    


#%% Calling Code
print("Initialising code . . .")
IW29=IW29_Automation_Class()

print("Opening/Connecting with SAP . . .")
IW29.connect_sap()
IW29.open_transaction()

print("Processing . . .")
IW29.process_transaction()

print("Done with Transaction, closing SAP . . .")
IW29._close_sap()
