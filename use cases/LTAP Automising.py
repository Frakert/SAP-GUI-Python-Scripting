import os
import sys

script_path = os.path.realpath(__file__)
parent_path = os.path.dirname(script_path)
main_folder, subfolder = os.path.split(parent_path)
src_folder = os.path.join(os.path.sep,main_folder,"Src")
sys.path.insert(0, src_folder)

from SAP_Automation_Class import SAP_Automation
import pandas as pd
from time import sleep

class LTAP_Automation_Class(SAP_Automation):

    def open_transaction(self):
        session=self.PRD_session
        session.SendCommand("/nIE02")
        return

    def process_transaction(self, equip_number: int) -> None:
        try:
            session=self.PRD_session

            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/usr/ctxtRM63E-EQUNR").text = equip_number
            session.findById("wnd[0]/usr/ctxtRM63E-EQUNR").caretPosition = 7
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[1]/btn[20]").press()
            session.findById("wnd[0]/usr/btn%#AUTOTEXT003").press()
            session.findById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0,1]").text = "LTAP"
            session.findById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0,1]").caretPosition = 4
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

        
        except Exception as err:
            self.logger.error(err)
            self.logger.info("Equipment errored was : %s"%(equip_number))
            #on error, close SAP and try again on the next equip
            self._close_sap()

            sleep(3)

            #Restart SAP
            self.connect_sap()
            self.open_transaction()

        return
    
    

#%%  Actual Code

list=pd.read_excel(r"C:\Users\klabbf\OneDrive - Canon Production Printing Netherlands B.V\Documents\SAP Scripting\LTAP_List.xlsx")
list=list.values.tolist()

print("Initialising code . . .")
LTAP=LTAP_Automation_Class()

print("Opening/Connecting with SAP . . .")
LTAP.connect_sap()
LTAP.open_transaction()

print("Processing . . .")
for equip in list:
    equip_number=equip[0]
    #equip_number=equip[0]
    print(equip)

    LTAP.process_transaction(equip_number) # Dynamic Dispatch method aka Overide method.

print("Done with Transaction, closing SAP . . .")
LTAP._close_sap()