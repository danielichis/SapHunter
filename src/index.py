from ast import Not
from processExtracts import process_xlsxFiles
from loadSap import startSAP, loadBankTemplates
from pathManagement import get_templates_path

def main():
    okExtracts=process_xlsxFiles() #procesamos todos los archivos en bruto y obtenemos sus plantillas
    if not(okExtracts):
        return
    sapInfo=get_templates_path() # con 4 ultimos digitos buscamos la info de cada plantilla para subir al sap
    #print(len(sapInfo))
    process=startSAP() #iniciamos sap
    for j,template in enumerate(sapInfo):
        #list to prove ["61539","42984"]
        #20210,70014 el extracto no esta disponible en memoria de datos bancarios
        #20635,66211 con error en el formato de las fechas
        #61539,42984 ok
        uatList=["20210"]
        if template["acountBin"] in uatList:
            try:
                loadBankTemplates(template) #cada template es un diccionario que tiene la ruta del archivo y la info de la cuenta
                print("Template loaded successfully")
                if j==len(sapInfo)-1:
                    print("Last template")
                    process.kill()
                else:
                    print("Next template")
            except Exception as e:
                print(e)
                if str(e).find("Sapgui Component") > 0:
                    print("ERROR AL ARRANCAR EL SAPGUI COMPONENT, REINICIANDO EL PROCESO...")
                    process.kill()
                    startSAP()
                print("Error al cargar el template")
                process.kill() #cerramos sap"""
            if j==len(sapInfo):
                print("Last template")
            else:
                print("Next template")
                    #startSAP() #iniciamos sap
if __name__ == "__main__":
    #process_xlsxFiles()
    #startSAP()
    main()