from locale import currency
from time import time
import openpyxl
import datetime
from datetime import datetime as dtime
from datetime import timedelta
from calendar import monthrange
import os
import pandas as pd
from pathManagement import get_current_path

def read_cuentas():
    configName="config.xlsx"
    newpath=os.path.join(get_current_path(),configName)
    cuentas=pd.read_excel(newpath,sheet_name="cuentas")
    return cuentas
def read_config():
    bankNamefile=os.path.join(get_current_path(),"config.xlsx")
    sheetName="bancos"
    wb = openpyxl.load_workbook(bankNamefile)
    sheet = wb[sheetName]
    n=sheet.max_row
    m=sheet.max_column
    listBanks=[x[0].value for x in sheet['A2:A'+str(n)]]
    listFields=[x.value for x in sheet['B1:P1'][0]]
    data={}
    for i,bank in enumerate(listBanks):
        data[bank]={}
        for j,field in enumerate(listFields):
            datacell=sheet.cell(row=i+2, column=j+2).value
            data[bank][field]=datacell
    return data
data=read_config()

def get_sheet(bankName,pathFile):
    print(data[bankName])
    sheetName=data[bankName]['NombreHoja']
    #print(bankNamefile,sheetName)
    wb = openpyxl.load_workbook(pathFile)
    sheet = wb[sheetName]
    return sheet
def get_dateRow(sheet,i,dateColumn):
    if sheet.cell(row=i, column=dateColumn).value==None:
        dateRow=''
    else:
        dateRow=sheet.cell(row=i, column=dateColumn).value
    return dateRow

def get_documentNr(bankName,sheet,i,documentColumn,codBancaColumn):
    #print("documento :")
    #print(sheet.cell(row=i, column=codBancaColumn).value)
    #print(type(sheet.cell(row=i, column=documentColumn).value))
    if bankName=='MERCANTIL':
        if sheet.cell(row=i, column=documentColumn).value==None or sheet.cell(row=i, column=documentColumn).value=='':
            if sheet.cell(row=i, column=codBancaColumn).value==None:
                nroDocument=''
                #print("acaaaa")
            else:
                nroDocument=sheet.cell(row=i, column=codBancaColumn).value
                #print("allaaaa")
        else:
            nroDocument=sheet.cell(row=i, column=documentColumn).value
            #print("aquiiiii")
    else:
        if sheet.cell(row=i, column=documentColumn).value==None:
            nroDocument=''
        else:
            nroDocument=sheet.cell(row=i, column=documentColumn).value
    return nroDocument
def get_description(bankName,sheet,i,descriptionColumn,nombreColum):
    if bankName=="MERCANTIL":
        if sheet.cell(row=i, column=descriptionColumn).value==None:
            description=''
        else:
            description=sheet.cell(row=i, column=descriptionColumn).value
        if sheet.cell(row=i, column=nombreColum).value==None:
            name=''
        else:
            name=sheet.cell(row=i, column=nombreColum).value
        description=name+"-"+description
    else:   
        if sheet.cell(row=i, column=descriptionColumn).value==None:
            description=''
        else:
            description=sheet.cell(row=i, column=descriptionColumn).value
    return description
def get_saldo(sheet,i,descriptionColumn,amountColumn):
    if sheet.cell(row=i, column=amountColumn).value==None:
        saldo=0
    else:
        saldo=float(str(sheet.cell(row=i, column=amountColumn).value).replace(",",""))
    return saldo
def get_amount(bankName,sheet,i,amountColumn,creditColumn,debitColumn):
    if bankName=="MERCANTIL" or bankName=="FASSIL":
        if sheet.cell(row=i, column=creditColumn).value==None:
            credit=0
        else:
            credit=float(str(sheet.cell(row=i, column=creditColumn).value).replace(",",""))

        if sheet.cell(row=i, column=debitColumn).value==None:
            debit=0
        else:
            debit=float(str(sheet.cell(row=i, column=debitColumn).value).replace(",",""))
        amount=credit-debit
    else:
        if sheet.cell(row=i, column=amountColumn).value==None:
            amount=0
        else:
            amount=float(str(sheet.cell(row=i, column=amountColumn).value).replace(",",""))
    return amount
def get_typetrx(description,amount):
    if description!='':
        if description.find("ITF")!=-1:
            typetrx="ZITF"
        else:
            if amount>0:
                typetrx="Z001"
            else:
                typetrx="Z002"
    else:
        if amount>0:
            typetrx="Z001"
        else:
            typetrx="Z002"
    return typetrx
def read_bank(fileMeta):
    bankName=fileMeta['bankName']
    sheetName=data[bankName]['NombreHoja']
    pathFile=fileMeta['path']
    sheet = get_sheet(bankName,pathFile)
    bankName=fileMeta['bankName']
    dataUnion=[]

    cuentaColumn=data[bankName]['CuentaCol']
    celdaFecha=data[bankName]['CeldaPeriodo']
    dateColumn=ord(celdaFecha[0])-64 #PARA SALIDA 1
    maxRowColumn=data[bankName]['MaxRowCol']
    documentColumn=data[bankName]['NroDocumCol'] #PARA SALIDA 2 

    creditColumn=data[bankName]['CreditCol']#OPCIONAL PARA SALIDA 3
    debitColumn=data[bankName]['DebitCol'] #OPCIONAL PARA SALIDA 3
    importColumn=data[bankName]['ImporteCol'] # ES PARA SALIDA 3
    saldoColumn=data[bankName]['SaldoCol'] #OPCIONAL PARA SALIDA 3
    nombreColum=data[bankName]['NombreCol']#OPCIONAL PARA SALIDA EN SOLO 1 BANCO 
    descripcionColumn=data[bankName]['DescripcionCol']#OPCIONAL ES PARA SALIDA 4
    codBancaColumn=data[bankName]['CodBancario']#OPCIONAL 
    
    dateformat=data[bankName]['FormatoFecha']
    dateExcel=sheet[celdaFecha].value
    dateCero = datetime.datetime.strptime(str(dateExcel), dateformat)
    print("------------",dateCero,"----------------")
    initialDate=datetime.datetime(dateCero.year,dateCero.month,1)
    finalDate=datetime.datetime(dateCero.year,dateCero.month,monthrange(dateCero.year,dateCero.month)[1])
    i=int(celdaFecha[1:])

    #print(i,sheet.max_row,descripcionColumn)
    while i<=sheet.max_row:
        if sheet.cell(row=i, column=maxRowColumn).value!=None:
            dateRow=get_dateRow(sheet,i,dateColumn)
            nroDocument=get_documentNr(bankName,sheet,i,documentColumn,codBancaColumn)
            saldo=get_saldo(sheet,i,saldoColumn,saldoColumn)
            description=get_description(bankName,sheet,i,descripcionColumn,nombreColum) #DINAMICO 
            amount=get_amount(bankName,sheet,i,importColumn,creditColumn,debitColumn) #DINAMICO
            typetrx=get_typetrx(description,amount)
            unionRow={
                "date":dateRow,
                "documentNr":nroDocument,
                "description":description,
                "type":typetrx,
                "amount":amount,
                "saldo":saldo,
            }
            dataUnion.append(unionRow)
        i=i+1
    initialBalance=dataUnion[0]['saldo']-dataUnion[0]['amount']
    finalBalance=dataUnion[len(dataUnion)-1]['saldo']
    #print(initialBalance,finalBalance)
    #print(pd.DataFrame(dataUnion))
    return {'data':dataUnion,'initialBalance':initialBalance,
    'finalBalance':finalBalance,'initialDate':initialDate,
    'finalDate':finalDate,'account':data[bankName]['CuentaCol'],'namComercial':data[bankName]['NombreComercial']}
def make_templates(infobank):
    binAccount=infobank['account']
    dateFname=dtime.today().strftime("%d%m%Y")
    fileName=f"{binAccount}-{dateFname}"
    extractFilePath=os.path.join(get_current_path(), "extractosBancarios",fileName)
    wb = openpyxl.load_workbook(extractFilePath)
    sheet = wb["UNION"]
    initialDate=infobank['initialDate'].strftime("%d/%m/%Y")
    finalDate=infobank['finalDate'].strftime("%d/%m/%Y")
    initialBalance=infobank['initialBalance']
    finalBalance=infobank['finalBalance']
    accountNumber=infobank['account']
    newPeriod=f"ESTADO DE CUENTA DEL {initialDate} AL {finalDate}"
    dataSap=infobank['data']
    binAccountouput=infobank['account'][:4]
    sheet["A1"]=newPeriod
    sheet["C4"]=accountNumber
    sheet["A9"]=initialBalance
    sheet["B9"]=finalBalance
    sheet["A8"]=initialDate
    sheet["B8"]=finalDate
    sheet["B3"]=infobank['namComercial']
    rowInit=17
    for i in range(len(dataSap)):
        sheet.cell(row=rowInit,column=1).value=dataSap[i]['date']
        sheet.cell(row=rowInit,column=2).value=dataSap[i]['documentNr']
        sheet.cell(row=rowInit,column=3).value=dataSap[i]['description']
        sheet.cell(row=rowInit,column=4).value=dataSap[i]['type']
        sheet.cell(row=rowInit,column=5).value=dataSap[i]['amount']
        rowInit=rowInit+1
    fileTemplateName=f"{binAccount}-{dateFname}.xlsx"
    fileTemplatePath=os.path.join(get_current_path(), "plantillSap",fileTemplateName)
    wb.save(fileTemplatePath)

def get_extrac_files():
    # get the current date
    tableAcounts=read_cuentas().values.tolist()
    datePathValue=datetime.date.today().strftime("%d%m%Y")
    newpath=os.path.join(get_current_path(), "extractosBancarios",datePathValue)
    files = os.listdir(newpath)
    paths=[]
    for f in files:
        for tb in tableAcounts:
            #print(str(tb[3])[-4:],"-----",f[:4])
            if f[:4]==str(tb[3])[-4:]:
                society=tb[4]
                codeBank=tb[5]
                currency=tb[6]
                bankName=tb[2]
                break
        filemeta={
            "name":f,
            "path":os.path.join(newpath,f),
            "society":society,
            "codeBank":codeBank,
            "currency":currency,
            "bankName":bankName
        }
                

        #print(f[:4])
        paths.append(filemeta)
    return paths


def process_xlsxFiles():
    pathFiles=get_extrac_files()
    print(pathFiles)
    for fileMeta in pathFiles:
        infobank=read_bank(fileMeta)
        make_templates(infobank)
#infobank=read_bank("ECONOMICO","banco economico.xlsx","ECONOMICO")
#make_templates(infobank)
if __name__ == "__main__":
    process_xlsxFiles()
    #print(data)
#process_xlsxFiles()

#print(df)