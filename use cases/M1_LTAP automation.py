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
import typing

class M1_LTAP_Automation_Class(SAP_Automation):

    def open_transaction(self) -> None:
        session=self.PRD_session
        session.SendCommand("/nIE02")
        return None

    def process_transaction(self, row: tuple) -> None:
        """ 
        Void function which fills in a few fields in the IE02 transaction.
        Fills in Equipment cost, currency tool creation date. And 5 LTAP (Z_Equipment) fields.
        
        Data = [SAP_Nr, new_date, price, price_currency, part_nr, recur_maint, eop, war_exp, eol]
        """
        try:
            session=self.PRD_session
            if SAP_Nr == 0 or SAP_Nr=="0":
                self.logger.error("Loop broken")
                self.logger.info("An equipment was skipped because it did not contain a SAP Equipment Number")
                print("An equipment was skipped because it did not contain a SAP Equipment Number")
                return
                

            session.findById("wnd[0]/usr/ctxtRM63E-EQUNR").text = row[0] # Sap_Nr
            session.findById("wnd[0]/usr/ctxtRM63E-EQUNR").caretPosition = 7
            session.findById("wnd[0]").sendVKey(0)

            print(row)
            
            # Pay attention that every \ should be replaced by \\ because Python.
            if row[2] != '0': session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102B:SAPLITO0:1021/txtITOB-ANSWT").text = row[2] # price
            if row[3] != '0': session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102B:SAPLITO0:1021/ctxtITOB-WAERS").text = row[3] # currency
            if row[1] != '0': session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102B:SAPLITO0:1021/ctxtITOB-ANSDT").text = row[1] # new_date
            session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102B:SAPLITO0:1021/ctxtITOB-ANSDT").setFocus()
            session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102B:SAPLITO0:1021/ctxtITOB-ANSDT").caretPosition = 10
            session.findById("wnd[0]/tbar[1]/btn[20]").press()
            if row[4] != '0': session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,2]").text = row[4] # part_nr
            if row[5] != '0': session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,6]").text = row[5] # recur_main
            if row[6] != '0': session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,7]").text = row[6] # eop (end of production)
            if row[7] != '0': session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,8]").text = row[7] # war_exp
            if row[8] != '0': session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,9]").text = row[8] # eol (end of life)
            session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,9]").setFocus()
            session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,9]").caretPosition = 0
            session.findById("wnd[0]/tbar[0]/btn[0]").press()
            session.findById("wnd[0]/tbar[1]/btn[32]").press()
            session.findById("wnd[0]/tbar[0]/btn[11]").press()

        
        except Exception as err:
            self.logger.error(err)
            self.logger.info("Equipment errored was : %s"%(row))
            #on error, close SAP and try again on the next equip
            self._close_sap()

            sleep(3)

            #Restart SAP
            self.connect_sap()
            self.open_transaction()

        return None
    
    

#%%  Actual Code

import pandas as pd

LTAP = pd.read_excel("Nout_Z_Equip.xlsx") # Read file im not giving because of company info
LTAP.fillna(0,inplace=True)
LTAP = LTAP.values.tolist()

print("Initialising code . . .")
M1_LTAP=M1_LTAP_Automation_Class()

print("Opening/Connecting with SAP . . .")
M1_LTAP.connect_sap()
M1_LTAP.open_transaction()

print("Processing . . .")

# For every line in excel file, get values, put them into the process_transaction method
for equip in LTAP:
    # Defenitions
    SAP_Nr = str(int(equip[14]))
    creation_date = equip[13]
    new_date = creation_date[8:10] + "." + creation_date[5:7] + '.' + creation_date[:4]
    price, price_currency = str(int(equip[19])), equip[20]
    part_nr = str(int(equip[21]))
    recur_maint = str(int(equip[25]))
    eop = str(int(equip[26]))
    war_exp = str(int(equip[27]))
    eol = str(int(equip[28]))

 

    Data= [SAP_Nr, new_date, price, price_currency, part_nr, recur_maint, eop, war_exp, eol]

    M1_LTAP.process_transaction(row = Data)


print("Done with Transaction, closing SAP . . .")
M1_LTAP._close_sap()