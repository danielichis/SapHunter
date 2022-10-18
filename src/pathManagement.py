from asyncio import DatagramTransport
from fileinput import filename
from pathlib import Path
import datetime
import pandas as pd
import os
def createFolder(path,force):
    try:
        if not os.path.exists(path):
            print("folder dont exits")
            if force:
                os.makedirs(path)
                print("folder created")
            return False
        else:
            print("Folder already exists")
            return True
    except OSError:
        print ('Error: Creating directory. ' +  path)
def get_current_path():
    currentSrcPath = os.getcwd()
    currentPath = Path(currentSrcPath)
    return currentPath 

def get_account_sap_info(binAccountPath):
    accountsInfo=pd.read_excel(os.path.join(get_current_path(),"config.xlsx"),sheet_name="cuentas").to_dict("records")
    #DataSap=[]
    for acountRow in accountsInfo:
        print(str(acountRow["NRO.CUENTA"])[-4:],binAccountPath)
        if str(acountRow["NRO.CUENTA"])[-4:]==str(binAccountPath):
            print("encontrado, terminando....")
            return acountRow
def get_login_info():
    loginInfo=pd.read_excel(os.path.join(get_current_path(),"config.xlsx"),sheet_name="LoginSap").to_dict("records")
    return loginInfo

def get_templates_path():
    currentDate=datetime.datetime.now().date().strftime("%d%m%Y")
    dir_path = os.path.join(get_current_path(), "plantillasSap",currentDate)
    sapInfo = []
    # Iterate directory
    for path in os.listdir(dir_path):
        # check if current path is a file
        if os.path.isfile(os.path.join(dir_path, path)):
            binAccountPath=path[:4]
            sapRow=get_account_sap_info(binAccountPath)
            if sapRow!=None:
                fiels={
                    "acountBin":path[:4],
                    "path":os.path.join(dir_path, path),
                    "CodeBank":sapRow["Banco propio (Campo de SAP)"],
                    "currency":sapRow["MONEDA"],

                }
            sapInfo.append(fiels)

    return sapInfo
print(get_login_info()[2]['VALOR'])

