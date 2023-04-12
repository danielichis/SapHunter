from pathlib import Path
from datetime import timedelta,datetime
import sys
import pandas as pd
import os
import openpyxl

def readTemplateSap(sapInfo):
    wb=openpyxl.load_workbook(sapInfo["path"])
    sh=wb.worksheets[0]
    initialBalance=float("{:.2f}".format(float(sh["A9"].value)))
    finalBalance=float("{:.2f}".format(float(sh["B9"].value)))
    
    bankNamefile=os.path.join(get_current_path(),"config.xlsx")
    sheetName="LoginSap"
    wb = openpyxl.load_workbook(bankNamefile)
    sheet = wb[sheetName]
    optionCase=sheet["B6"].value
    
    with open(sapInfo["AuzugTxtPath"], 'r') as file:
        line=file.readlines()[0]
        initialAuzug = line.split(";")[5].strip()
        finalAuzug = line.split(";")[8].strip()

        if initialAuzug.find("-")>0:
            initialAuzug=initialAuzug.replace("-","")
            initialAuzug=-float(initialAuzug)

        if finalAuzug.find("-")>0:
            finalAuzug=finalAuzug.replace("-","")
            finalAuzug=-float(finalAuzug)
    initialAuzug=float("{:.2f}".format(float(initialAuzug)))
    finalAuzug=float("{:.2f}".format(float(finalAuzug)))
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
def createFolder(path,force,delete):
    try:
        if not os.path.exists(path):
            #print("folder doesn't exist")
            if force:
                os.makedirs(path)
                print("folder created")
            return False
        else:
            if delete:
                print("Folder already exists")
                files=os.listdir(path)
                for file in files:
                    if file.endswith(".xlsx"):
                        pass
                        #os.remove(path+"\\"+file)
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
def isNaN(num):
    return num!= num
def get_login_info():
    global loginInfo
    loginInfo=pd.read_excel(os.path.join(get_current_path(),"config.xlsx"),sheet_name="LoginSap").to_dict("records")

    return loginInfo
def testnulldate():
    cddte=get_login_info()[5]['VALOR']
    today=datetime.today().date()
    yesterday=today-timedelta(days=1)
    currentDateSap=yesterday.strftime("%d.%m.%Y")
    if isNaN(cddte):
        pass
    else:
        currentDateSap=datetime.strptime(str(cddte),"%Y-%m-%d %H:%M:%S").strftime("%d.%m.%Y")
    print(currentDateSap)

def get_templates_path():
    cddte=get_login_info()[5]['VALOR']
    today=datetime.today().date()
    yesterday=today-timedelta(days=1)
    currentday=today-timedelta(days=1)
    currentDateSap=yesterday.strftime("%d.%m.%Y")
    currentDateFolder=currentday.strftime("%d%m%Y")
    if isNaN(cddte):
        pass
    else:
        currentDateSap=datetime.strptime(str(cddte),"%Y-%m-%d %H:%M:%S").strftime("%d.%m.%Y")

    dir_path = os.path.join(get_current_path(), "plantillasSap",currentDateFolder)
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
                        "currentDate":currentDateSap
                    }
                sapInfo.append(fiels)

    return sapInfo
print(testnulldate())

