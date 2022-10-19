from os import environ
import win32com.client
import sys
import subprocess
import time

def startSAP(environment):
    global command2, sapGuiAuto, application, connection, session
    command2 =r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    proc = subprocess.Popen([command2, '-new-tab'])
    time.sleep(2)

    sapGuiAuto = win32com.client.GetObject('SAPGUI')
    if not type(sapGuiAuto) == win32com.client.CDispatch:
        pass

    application = sapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        sapGuiAuto = None
        pass

    #connection = application.OpenConnection("SAP QAS", True)
    connection = application.OpenConnection(environment, True)
    if not type(connection) == win32com.client.CDispatch:
        application = None
        sapGuiAuto = None
        pass

    session = connection.Children(0)
    #print(help(connection))
    if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        sapGuiAuto = None
        pass

def loadBankTemplates():

    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "BOT"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "botpruebas**"
    session.findById("wnd[0]").sendVKey(0)

    # Ingresar datos del banco
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZFI_EXTBAN"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/usr/ctxtFECHCONT").text = "01.05.2022"
    session.findById("wnd[0]/usr/ctxtSOCIEDAD").text = "GV01"
    session.findById("wnd[0]/usr/ctxtBANCOID").text = "BMS06"
    session.findById("wnd[0]/usr/ctxtCTAID").text = "42984"
    session.findById("wnd[0]/usr/ctxtMONEDA").text = "BOB"
   
    session.findById("wnd[0]/usr/txtSINI").text = "0"
    session.findById("wnd[0]/usr/ctxtARCHIVO").setFocus
    path_xlsx=r"C:\Users\Administrador.WIN-C8USBNGG6F4\OneDrive - industrias venado\Documentos\2984-21062022.xlsx"
    session.findById("wnd[0]/usr/ctxtARCHIVO").text = path_xlsx
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(1)
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    time.sleep(1)
    session.findById("wnd[1]/usr/ctxtRLGRAP-FILENAME").text =r"C:\Users\Administrador.WIN-C8USBNGG6F4\OneDrive - industrias venado\Documentos\sapProyecto\auszug.txt"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(1)
    session.findById("wnd[1]/usr/ctxtRLGRAP-FILENAME").text =r"C:\Users\Administrador.WIN-C8USBNGG6F4\OneDrive - industrias venado\Documentos\sapProyecto\umsatz.txt"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(1)
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    time.sleep(1)
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00116")
    time.sleep(1)
    session.findById("wnd[0]/usr/chkEINLESEN").selected = "true"
    session.findById("wnd[0]/usr/ctxtAUSZFILE").text =r"C:\Users\Administrador.WIN-C8USBNGG6F4\OneDrive - industrias venado\Documentos\sapProyecto\auszug.txt"
    session.findById("wnd[0]/usr/ctxtUMSFILE").text =r"C:\Users\Administrador.WIN-C8USBNGG6F4\OneDrive - industrias venado\Documentos\sapProyecto\umsatz.txt"
    session.findById("wnd[0]/usr/radPA_TEST").selected = "true"
    session.findById("wnd[0]/usr/chkP_KOAUSZ").selected = "true"
    session.findById("wnd[0]/usr/chkP_BUPRO").selected = "true"
    session.findById("wnd[0]/usr/chkP_STATIK").selected = "true"
    session.findById("wnd[0]/usr/chkPA_LSEPA").selected = "true"
    session.findById("wnd[0]/usr/radPA_TEST").setFocus
    session.findById("wnd[0]/usr/radPA_TEST").setFocus
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode("F00113")
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00115"
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").topNode = "Favo"
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00115")
    time.sleep(0.5)

    session.findById("wnd[1]/usr/ctxtSL_BUKRS-LOW").text = "GV01"
    session.findById("wnd[1]/usr/ctxtSL_HBKID-LOW").text = "BMS06"
    session.findById("wnd[1]/usr/ctxtSL_HKTID-LOW").text = "42984"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    p=session.findById("wnd[0]/shellcont/shell").getNodeChildrenCount("0101")
    session.findById("wnd[0]/shellcont/shell").expandNode(f"01010{p}")
    session.findById("wnd[0]/shellcont/shell").selectedNode = (f"01010{p}0001")
    session.findById("wnd[0]/shellcont/shell").nodeContextMenu(f"01010{p}0001")
    session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("BS_POST_ITEMS")
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

# def processInSAP():
#     path_xlsx=r"C:\Users\crist\OneDrive - UNIVERSIDAD NACIONAL DE INGENIERIA\Venado\Cris\2984-21062022.xlsx"
#     environment = "QAS - EHP8 on HANA"

if __name__=='__main__':
    environment= "QAS - EHP8 on HANA"
    startSAP(environment)
    loadBankTemplates()





# lastOne=False
# i=1
# while lastOne == False:
#     e=str(i).zfill(3)
#     try:
#         session.findById("wnd[0]/shellcont/shell").selectedNode = f"01010{e}"
#         print(e)
        
#     except:
#         lastOne=True
#         print("chacón hdp")
#     i=i+1
# e=str(i-2).zfill(3)