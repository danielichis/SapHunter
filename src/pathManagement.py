from pathlib import Path
import datetime
import sys
import pandas as pd
import os
import openpyxl

def readTemplateSap(sapInfo):
    wb=openpyxl.load_workbook(sapInfo["path"])
    sh=wb.worksheets[0]
    initialBalance=sh["A9"].value
    finalBalance=sh["B9"].value

    bankNamefile=os.path.join(get_current_path(),"config.xlsx")
    sheetName="LoginSap"
    wb = openpyxl.load_workbook(bankNamefile)
    sheet = wb[sheetName]
    optionCase=sheet["B6"].value
    
    with open(sapInfo["AuzugTxtPath"], 'r') as file:
        line=file.readlines()[0]
        initialAuzug = line.split(";")[5]
        finalAuzug = line.split(";")[8]

        if initialAuzug.find("-")>0:
            initialAuzug=initialAuzug.replace("-","")
            initialAuzug=-float(initialAuzug)

        if finalAuzug.find("-")>0:
            finalAuzug=finalAuzug.replace("-","")
            finalAuzug=-float(finalAuzug)
            
    print(initialBalance,finalBalance,initialAuzug,finalAuzug)
    if optionCase=="Doble":
        if initialBalance==initialAuzug and finalBalance==finalAuzug:
            return True
        else:
            return False
    elif optionCase=="Simple":
        if initialBalance==finalAuzug:
            return True
        else:
            return False
    else:
        return True
def createFolder(path,force):
    try:
        if not os.path.exists(path):
            #print("folder doesn't exist")
            if force:
                os.makedirs(path)
                print("folder created")
            return False
        else:
            print("Folder already exists")
            return True
    except OSError:
        print ('Error: Creating directory. ' +  path)
def get_current_path2():
    currentSrcPath = os.getcwd()
    print(currentSrcPath)
    currentPath = Path(currentSrcPath)
    return currentPath

def get_current_path():
    config_name = 'myapp.cfg'
    # determine if application is a script file or frozen exe
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)
    application_path2 = Path(application_path)
    return application_path2.parent.absolute()
def delete_txtFiles(txtPath):
    for path in os.listdir(txtPath):
        # check if current path is a file
        if os.path.isfile(os.path.join(txtPath, path)):
            if path[-4:]==".txt":
                if path=="auszug.txt" or path=="umsatz.txt":
                    os.remove(os.path.join(txtPath, path))
                    print("txt file deleted")

def get_account_sap_info(binAccountPath):
    accountsInfo=pd.read_excel(os.path.join(get_current_path(),"config.xlsx"),sheet_name="cuentas").to_dict("records")
    #DataSap=[]
    for acountRow in accountsInfo:
        #print(str(acountRow["NRO.CUENTA"])[-4:],binAccountPath)
        if str(acountRow["NRO.CUENTA"])[-4:]==str(binAccountPath):
            #print("encontrado, terminando....")
            return acountRow
def get_login_info():
    loginInfo=pd.read_excel(os.path.join(get_current_path(),"config.xlsx"),sheet_name="LoginSap").to_dict("records")
    return loginInfo

def get_templates_path():
    currentDate=datetime.datetime.now().date().strftime("%d%m%Y")
    currentDateSap=datetime.datetime.now().date().strftime("%d.%m.%Y")
    dir_path = os.path.join(get_current_path(), "plantillasSap",currentDate)
    sapInfo = []
    # Iterate directory
    for path in os.listdir(dir_path):
        # check if current path is a file
        if os.path.isfile(os.path.join(dir_path, path)):
            if path[-5:]==".xlsx":
                binAccountPath=path[:4]
                sapRow=get_account_sap_info(binAccountPath)
                if sapRow!=None:
                    txtNameAuzug=f"auszug.txt"
                    AzugPath=os.path.join(dir_path,txtNameAuzug)
                    txtNameUmzat=f"umsatz.txt"
                    UmzatPath=os.path.join(dir_path,txtNameUmzat)
                    fiels={
                        "acountBin":str(sapRow["NRO.CUENTA"])[-5:],
                        "path":os.path.join(dir_path, path),
                        "AuzugTxtPath":AzugPath,
                        "umzatTxtPath":UmzatPath,
                        "folderPath":dir_path,
                        "societyCode":sapRow["Sociedad (Campo de SAP)"],
                        "CodeBank":sapRow["Banco propio (Campo de SAP)"],
                        "currency":sapRow["MONEDA"],
                        "abrCurrency":sapRow["ABR MONEDA"],
                        "currentDate":"01.05.2022"
                    }
                sapInfo.append(fiels)

    return sapInfo
print(get_current_path())

