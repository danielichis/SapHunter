from processExtracts import process_xlsxFiles, write_log
from loadSap import startSAP, loadBankTemplates
from pathManagement import get_templates_path,get_login_info

def main():
    runnerMode=get_login_info()[6]['VALOR']
    if runnerMode=="PLANTILLAS" or runnerMode=="COMPLETO":
        okExtracts=process_xlsxFiles() #procesamos todos los archivos en bruto y obtenemos sus plantillas 1era parte
        if not(okExtracts):
            return
    if runnerMode=="CARGA_SAP" or runnerMode=="COMPLETO":
        sapInfo=get_templates_path() # con 4 ultimos digitos buscamos la info de cada plantilla para subir al sap
        process=startSAP() #iniciamos sap
        for j,template in enumerate(sapInfo):
            print(f"procesando plantilla {j+1} de {len(sapInfo)}")
            try:
                write_log("","CARGANDO CUENTA: "+template["acountBin"],template["path"])
                loadBankTemplates(template) #cada template es un diccionario que tiene la ruta del archivo y la info de la cuenta
                write_log(" ","CARGADO CORRECTAMENTE ",template["path"])
            except Exception as e:
                write_log(" ",e,template["path"])
            write_log("","\n",template["path"])
        try:
            process.kill() #cerramos sap"""
            process.kill() #cerramos sap"""
            process.kill() #cerramos sap""" 
        except:
            pass
if __name__ == "__main__":
    #process_xlsxFiles()
    #startSAP()
    main()