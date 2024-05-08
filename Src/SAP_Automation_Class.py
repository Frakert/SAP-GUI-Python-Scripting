""""
Made by: KLABBF (Freek Klabbers)
Date: 13-12-2023

I Created this code as a baseline for SAP automation classes.
"""

#%% Imports & Definitions
import win32com.client
import subprocess 
from time import sleep
from getpass import getuser
import logging
import time
#%% Building a Class

class SAP_Automation:
  """"
  Made by: KLABBF (Freek Klabbers)
  Date: 13-12-2023

  A Simple Class that will funcion as a parent class for SAP Automation tasks.
  It will manage the SAP instance and its subwindows and offer standard opperations that could be peformed.

  Dependencies: win32com.client, subprocess, time, getpass, loggingn
  """

  def __init__(self) -> None:
     """Initialse the class, adds a debug log to the downloads folder"""
     self.user=getuser()
     now = time.strftime("%Y-%m-%d_%H_%M_%S")
     logging.basicConfig(filename='C:/Users/'+ self.user + '/Downloads/LTAP' + now + '.log', level=logging.DEBUG,format='%(asctime)s %(levelname)s %(name)s %(message)s')
     self.logger=logging.getLogger(__name__)
     return
  

  def _open_sap(self) -> None:
      """SAP needs to be open for most commands to work, so just open it. Give the PC some time to do so."""

      path = 'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'
      subprocess.Popen(path)
      sleep(3)
      return
  
  def _close_sap(self) -> None:
     """Close SAP using taskmanager kill."""
     subprocess.call("taskkill /f /im saplogon.exe", shell=True)
     return

  def connect_sap(self) -> None:
     """
     Openup SAP, Connect to the scripting Engine and use it to open up SAP PRD with Single Sign On (SSO).
     """

     self._open_sap()
     self.__sap_app = win32com.client.GetObject("SAPGUI").GetScriptingEngine
     self.__sap_app.OpenConnection("PRD - ERP Production (SSO)", True)

     self.__gui_session = self.__sap_app.Children(0) # References the first window
     self.PRD_session = self.__gui_session.Children(0) # References the second window
     return

  def process_transaction(self) -> None:
     """
     Made as a Override method.
     """

     session=self.PRD_session

     session.SendCommand("/nIW29")

     # Select Transaction paramaterss 
     session.findById("wnd[0]").resizeWorkingPane(210,28,"false")
     session.findById("wnd[0]/usr/chkDY_MAB").selected = "true" # include completed transactions
     session.findById("wnd[0]/usr/ctxtDATUV").text = "1.1.2023" # begin date
     session.findById("wnd[0]/usr/ctxtDATUV").setFocus()
     session.findById("wnd[0]/usr/ctxtDATUV").caretPosition = 3
     session.findById("wnd[0]/tbar[1]/btn[8]").press()
     return
  
  def export_to_excel(self) -> None:
     """
     Export the transaction made to Excel.
     """
     session=self.PRD_session

     session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
     session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
     session.findById("wnd[1]/tbar[0]/btn[0]").press()
     session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
     session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 8
     session.findById("wnd[1]").sendVKey(4)
     session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"C:\Users\{}\Downloads".format(self.user)
     session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
     session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 12
     session.findById("wnd[2]/tbar[0]/btn[0]").press()
     session.findById("wnd[1]/tbar[0]/btn[0]").press()

     # Wait a bit and then kill the Excel window opening
     sleep(3)
     subprocess.call("taskkill /f /im EXCEL.EXE", shell=True)
     return


  def __exit__(self) -> None:
     self._close_sap()