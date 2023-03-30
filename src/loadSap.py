from cmath import exp
from concurrent.futures import process
import win32com.client
import subprocess
import time
from pathManagement import delete_txtFiles, get_login_info,readTemplateSap
from processExtracts import write_log
#llerena vago ctmr

def startSAP():
    loginData=get_login_info()
    user=str(get_login_info()[0]['VALOR'])
    password=str(get_login_info()[1]['VALOR'])
    enviroment=str(get_login_info()[2]['VALOR'])
    pathSap=str(get_login_info()[3]['VALOR'])
    global command2, sapGuiAuto, application, connection, session
    command2 =pathSap
    proc = subprocess.Popen([command2, '-new-tab'])
    time.sleep(2)
    try:
        sapGuiAuto = win32com.client.GetObject('SAPGUI')
    except:
        print("Error al iniciar el SAPGUI Component")
        proc.kill()
        proc=startSAP()
        return proc
    if not type(sapGuiAuto) == win32com.client.CDispatch:
        pass

    application = sapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        sapGuiAuto = None
        pass

    #connection = application.OpenConnection("SAP QAS", True)
    connection = application.OpenConnection(enviroment, True)
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
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
    session.findById("wnd[0]").sendVKey(0)
    print("SAP STARTED SUCCESSFULLY...")
    return proc

def make_num(p):
    l=len(str(p))
    num="01010000"
    s=num[:8-l]+str(p)
    return s

def loadBankTemplates(infoSap):
    # Ingresar datos del banco
    delete_txtFiles(infoSap["folderPath"])
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZFI_EXTBAN"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/usr/ctxtFECHCONT").text = infoSap["currentDate"]
    session.findById("wnd[0]/usr/ctxtSOCIEDAD").text = infoSap["societyCode"]
    session.findById("wnd[0]/usr/ctxtBANCOID").text = infoSap["CodeBank"]
    session.findById("wnd[0]/usr/ctxtCTAID").text = infoSap["acountBin"]
    session.findById("wnd[0]/usr/ctxtMONEDA").text = infoSap["abrCurrency"]
   
    session.findById("wnd[0]/usr/txtSINI").text = "0"
    session.findById("wnd[0]/usr/ctxtARCHIVO").setFocus
    session.findById("wnd[0]/usr/ctxtARCHIVO").text = infoSap["path"]
    rutaExcel=infoSap["path"]
    print(f"Template loaded successfully{rutaExcel}")
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(1)
    try:
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    except:
        session.endTransaction()
        time.sleep(2)
        raise Exception("ERROR DE FORMATO DE ARCHIVO EXCEL")
    time.sleep(1)
    session.findById("wnd[1]/usr/ctxtRLGRAP-FILENAME").text =infoSap["AuzugTxtPath"]
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(1)
    session.findById("wnd[1]/usr/ctxtRLGRAP-FILENAME").text =infoSap["umzatTxtPath"]
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(1)
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    time.sleep(1)
    #HACER LA VALIDACION DEL SALDO INICIAL,SALDO FINAL DEL TXT CON EL EXCEL. LA INFO DEL EXCEL VIENE EN EL DICCIONARIO
    # QUE DEBEMOS AGREGAR COMO NUEVO PARAMETRO EN ESTA FUNCION 
    #UNA FUNCION PARA LEER EL SALDO INICIAL Y EL SALDO FINAL DEL TXT Y VALIDARLOS CON EL EXCEL
    rt=readTemplateSap(infoSap)
    if rt==True:
        pass
    else:
        #write_log(" ","ERROR EN LA VALIDACION DE TXT-EXCEL",infoSap["path"])
        session.endTransaction()
        time.sleep(2)
        raise Exception("ERROR EN VALIDACION DE SALDOS")
    #session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00116")
    session.findById("wnd[0]/tbar[0]/okcd").text = "FF_5"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)
    session.findById("wnd[0]/usr/chkEINLESEN").selected = "true"
    session.findById("wnd[0]/usr/ctxtAUSZFILE").text =infoSap["AuzugTxtPath"]
    session.findById("wnd[0]/usr/ctxtUMSFILE").text =infoSap["umzatTxtPath"]
    session.findById("wnd[0]/usr/radPA_TEST").selected = "true"
    session.findById("wnd[0]/usr/chkP_KOAUSZ").selected = "true"
    session.findById("wnd[0]/usr/chkP_BUPRO").selected = "true"
    session.findById("wnd[0]/usr/chkP_STATIK").selected = "true"
    session.findById("wnd[0]/usr/chkPA_LSEPA").selected = "true"

    session.findById("wnd[0]/usr/radPA_TEST").setFocus
    session.findById("wnd[0]/usr/radPA_TEST").setFocus
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except:
        pass
    session.endTransaction()

    # session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00115"
    # session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").topNode = "Favo"
    # session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00115")
    
    session.findById("wnd[0]/tbar[0]/okcd").text = "FEBAN"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    session.findById("wnd[1]/usr/ctxtSL_BUKRS-LOW").text = infoSap["societyCode"]
    session.findById("wnd[1]/usr/ctxtSL_HBKID-LOW").text = infoSap["CodeBank"]
    session.findById("wnd[1]/usr/ctxtSL_HKTID-LOW").text = infoSap["acountBin"]
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    p=session.findById("wnd[0]/shellcont/shell").getNodeChildrenCount("0101")
    p=make_num(p)
    session.findById("wnd[0]/shellcont/shell").expandNode(p)
    session.findById("wnd[0]/shellcont/shell").selectedNode = (p)
    session.findById("wnd[0]/shellcont/shell").nodeContextMenu(p)
    session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("BS_POST_ITEMS")
    text=session.findById("wnd[0]/sbar/pane[0]").text

    if text.find("se contabilizó")>-1:
        session.endTransaction()
        pass
    else:
        session.endTransaction()
        session.endTransaction()
        session.endTransaction()
        raise Exception(text)