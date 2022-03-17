

import tkinter as tk
import win32com.client
import sys
import subprocess
import time
import openpyxl
from pathlib import Path

xlsx_file = Path('zxc.xlsx')

wb_obj = openpyxl.load_workbook(xlsx_file)
sheet = wb_obj.active
username = sheet["A1"].value
password = sheet["B1"].value

def sap():
    try:

        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(10)

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return
        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        connection = application.OpenConnection('S48 [IN-BLR-1709]', True)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return
        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return
        session.findById("wnd[0]").maximize()

        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = 600
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus()
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 9
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        #session.findById("wnd[0]/tbar[0]/okcd").text = "se11"
        #session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "se24"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/ctxtSEOCLASS-CLSNAME").text = "ztest_class46134"
        session.findById("wnd[0]/usr/ctxtSEOCLASS-CLSNAME").caretPosition = 11
        session.findById("wnd[0]/usr/btnPUSH_CREATE").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/txtVSEOCLASS-DESCRIPT").text = "add class"
        session.findById("wnd[1]/usr/txtVSEOCLASS-DESCRIPT").setFocus()
        session.findById("wnd[1]/usr/txtVSEOCLASS-DESCRIPT").caretPosition = 9
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtKO007-L_DEVCLASS").text = "z46126613pkg"
        session.findById("wnd[1]/usr/ctxtKO007-L_DEVCLASS").caretPosition = 12
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]").sendVKey(4)

        session.findById("wnd[2]/usr/lbl[33,16]").setFocus()
        session.findById("wnd[2]/usr/lbl[33,16]").caretPosition = 9
        session.findById("wnd[2]").sendVKey(2)

        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT").select()
        session.findById("wnd[0]").resizeWorkingPane(69, 23, False)

        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/txtDY_0252-CPDNAME[0,0]").text = "input1"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/txtDY_0252-CPDNAME[0,1]").text = "input2"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/ctxtDY_0252-CHA_ATTDECLTYP[1,0]").text = "instance a"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/ctxtDY_0252-CHA_ATTDECLTYP[1,0]").setFocus()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/ctxtDY_0252-CHA_ATTDECLTYP[1,0]").caretPosition = 10
        session.findById("wnd[0]").sendVKey(4)

        session.findById("wnd[1]/usr/lbl[5,3]").setFocus()
        session.findById("wnd[1]/usr/lbl[5,3]").caretPosition = 1
        session.findById("wnd[1]").sendVKey(2)

        session.findById(
            "wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/ctxtDY_0252-CHA_ATTDECLTYP[1,1]").setFocus()
        session.findById(
            "wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/ctxtDY_0252-CHA_ATTDECLTYP[1,1]").caretPosition = 0
        session.findById("wnd[0]").sendVKey( 4)

        session.findById("wnd[1]/usr/lbl[5,3]").setFocus()
        session.findById("wnd[1]/usr/lbl[5,3]").caretPosition = 3
        session.findById("wnd[1]").sendVKey(2)

        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/ctxtDY_0252-CHA_EXPOSURE[2,0]").text = "p"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/ctxtDY_0252-CHA_EXPOSURE[2,0]").setFocus()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/ctxtDY_0252-CHA_EXPOSURE[2,0]").caretPosition = 1
        session.findById("wnd[0]").sendVKey(4)

        session.findById("wnd[1]/usr/lbl[5,5]").setFocus()
        session.findById("wnd[1]/usr/lbl[5,5]").caretPosition = 2
        session.findById("wnd[1]").sendVKey(2)

        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/ctxtDY_0252-CHA_EXPOSURE[2,1]").setFocus()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/ctxtDY_0252-CHA_EXPOSURE[2,1]").caretPosition = 0
        session.findById("wnd[0]").sendVKey(4)

        session.findById("wnd[1]/usr/lbl[5,5]").setFocus()
        session.findById("wnd[1]/usr/lbl[5,5]").caretPosition = 5
        session.findById("wnd[1]").sendVKey(2)

        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/ctxtVSEOATTRIB-TYPE[5,0]").text = "N"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/ctxtVSEOATTRIB-TYPE[5,1]").text = "N"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/btnDY_0252-PUSH_MORE[6,0]").setFocus()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_ATT/ssubCSS:SAPLSEOD:0252/tblSAPLSEODAC/btnDY_0252-PUSH_MORE[6,0]").press()
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        session.findById("wnd[0]/tbar[1]/btn[18]").press()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD").select()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0253/tblSAPLSEODMC/txtDY_0253-CPDNAME[0,0]").text = "addBoth"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0253/tblSAPLSEODMC/ctxtDY_0253-CHA_MTDDECLTYP[1,0]").text = "in"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0253/tblSAPLSEODMC").columns.elementAt(2).width = 9
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0253/tblSAPLSEODMC/ctxtDY_0253-CHA_MTDDECLTYP[1,0]").setFocus()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0253/tblSAPLSEODMC/ctxtDY_0253-CHA_MTDDECLTYP[1,0]").caretPosition = 2
        session.findById("wnd[0]").sendVKey(4)

        session.findById("wnd[1]/usr/lbl[5,3]").setFocus()
        session.findById("wnd[1]/usr/lbl[5,3]").caretPosition = 5
        session.findById("wnd[1]").sendVKey(2)

        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0253/tblSAPLSEODMC/ctxtDY_0253-CHA_EXPOSURE[2,0]").text = "Pub"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0253/tblSAPLSEODMC/ctxtDY_0253-CHA_EXPOSURE[2,0]").setFocus()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0253/tblSAPLSEODMC/ctxtDY_0253-CHA_EXPOSURE[2,0]").caretPosition = 3
        session.findById("wnd[0]").sendVKey(4)

        session.findById("wnd[1]/usr/lbl[5,3]").setFocus()
        session.findById("wnd[1]/usr/lbl[5,3]").caretPosition = 6
        session.findById("wnd[1]").sendVKey(2)

        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0253/tblSAPLSEODMC/txtDY_0253-CPDNAME[0,0]").setFocus()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0253/tblSAPLSEODMC/txtDY_0253-CPDNAME[0,0]").caretPosition = 6
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0253/btnPUSH_PARAMETERS").press()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0352/tblSAPLSEODPC/txtVSEOMEPARA-SCONAME[0,0]").text = "im_input1"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0352/tblSAPLSEODPC/txtVSEOMEPARA-SCONAME[0,1]").text = "im_input2"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0352/tblSAPLSEODPC/ctxtDY_0352-CHA_PARDECLTYP[1,0]").text = "Importin"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0352/tblSAPLSEODPC/ctxtDY_0352-CHA_PARDECLTYP[1,0]").setFocus()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0352/tblSAPLSEODPC/ctxtDY_0352-CHA_PARDECLTYP[1,0]").caretPosition = 8
        session.findById("wnd[0]").sendVKey(4)

        session.findById("wnd[1]/usr/lbl[5,3]").setFocus()
        session.findById("wnd[1]/usr/lbl[5,3]").caretPosition = 5
        session.findById("wnd[1]").sendVKey(2)

        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0352/tblSAPLSEODPC").columns.elementAt(1).width = 8
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0352/tblSAPLSEODPC/ctxtDY_0352-CHA_PARDECLTYP[1,1]").setFocus()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0352/tblSAPLSEODPC/ctxtDY_0352-CHA_PARDECLTYP[1,1]").caretPosition = 0
        session.findById("wnd[0]").sendVKey(4)

        session.findById("wnd[1]/usr/lbl[5,3]").setFocus()
        session.findById("wnd[1]/usr/lbl[5,3]").caretPosition = 9
        session.findById("wnd[1]").sendVKey(2)

        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0352/tblSAPLSEODPC/ctxtVSEOMEPARA-TYPE[5,0]").text = "N"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0352/tblSAPLSEODPC/ctxtVSEOMEPARA-TYPE[5,1]").text = "N"
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0352/tblSAPLSEODPC/txtVSEOMEPARA-PARVALUE[6,0]").setFocus()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0352/tblSAPLSEODPC/txtVSEOMEPARA-PARVALUE[6,0]").caretPosition = 0
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0352/btnPUSH_BACK").press()
        session.findById("wnd[0]/usr/tabsCTS/tabpTAB_MTD/ssubCSS:SAPLSEOD:0253/tblSAPLSEODMC/txtDY_0253-CPDNAME[0,0]").caretPosition = 3
        session.findById("wnd[0]").sendVKey(2)

        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("" + vbCr + "" + vbLf + "", 1, 18)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("    ", 2, 1)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("i", 2, 5)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("n", 2, 6)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("p", 2, 7)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("u", 2, 8)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("t", 2, 9)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("1", 2, 10)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText(" ", 2, 11)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("=", 2, 12)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText(" ", 2, 13)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("i", 2, 14)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("m", 2, 15)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("_", 2, 16)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("i", 2, 17)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("n", 2, 18)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("p", 2, 19)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("u", 2, 20)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("t", 2, 21)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("1", 2, 22)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText(".", 2, 23)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("" + vbCr + "" + vbLf + "", 2, 24)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("    ", 3, 1)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("i", 3, 5)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("n", 3, 6)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("p", 3, 7)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("u", 3, 8)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("t", 3, 9)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("2", 3, 10)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText(" ", 3, 11)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("=", 3, 12)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText(" ", 3, 13)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("i", 3, 14)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("m", 3, 15)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("_", 3, 16)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("i", 3, 17)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("n", 3, 18)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("p", 3, 19)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("u", 3, 20)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("t", 3, 21)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText("2", 3, 22)
        session.findById("wnd[0]/usr/subEDITORSUBSCREEN:SAPLEDITOR_START:8430/cntlEDITOR/shellcont/shell").insertText(".", 3, 23)
        session.findById("wnd[0]").sendVKey(11)
        session.findById("wnd[0]/tbar[1]/btn[18]").press

    except Exception as e:
        print("Error in Execution",e)



    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None

sap()
