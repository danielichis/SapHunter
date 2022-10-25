from ast import Not
from processExtracts import process_xlsxFiles, write_log
from loadSap import startSAP, loadBankTemplates
from pathManagement import get_templates_path

def main():
    okExtracts=process_xlsxFiles() #procesamos todos los archivos en bruto y obtenemos sus plantillas 1era parte
    if not(okExtracts):
        return
    sapInfo=get_templates_path() # con 4 ultimos digitos buscamos la info de cada plantilla para subir al sap
    process=startSAP() #iniciamos sap
    for j,template in enumerate(sapInfo):
        print(f"procesando plantilla {j+1} de {len(sapInfo)}")
        #list to prove ["61539","42984"]
        #20210,70014 el extracto no esta disponible en memoriad de datos bancarios
        #20635,66211 con error en el formato de las fechas
        #61539,42984 okok
        #uatList=["20210"]
        #if template["acountBin"] in uatList:
        try:
            write_log("","CARGANDO AL SAP: "+template["acountBin"],template["path"])
            loadBankTemplates(template) #cada template es un diccionario que tiene la ruta del archivo y la info de la cuenta
            write_log(" ","CARGADO CORRECTAMENTE",template["path"])
        except Exception as e:
            write_log(" ",e,template["path"])
        write_log("","\n",template["path"])
    process.kill() #cerramos sap"""
if __name__ == "__main__":
    #process_xlsxFiles()
    #startSAP()
    main()