from processExtracts import process_xlsxFiles
from loadSap import startSAP, loadBankTemplates
from pathManagement import get_templates_path

def main():
    process_xlsxFiles() #procesamos todos los archivos en bruto y obtenemos sus plantillas
    sapInfo=get_templates_path() # con 4 ultimos digitos buscamos la info de cada plantilla para subir al sap
    #print(len(sapInfo))
    process=startSAP() #iniciamos sap
    for j,template in enumerate(sapInfo):
        #print(template)
        if template["acountBin"]=="42984":
            try:
                loadBankTemplates(template) #cada template es un diccionario que tiene la ruta del archivo y la info de la cuenta
                print("Template loaded successfully")
            except:
                print("Error al cargar el template")
                process.kill() #cerramos sap
                if j==len(sapInfo):
                    print("Last template")
                else:
                    print("Next template")
                    #startSAP() #iniciamos sap
if __name__ == "__main__":
    #process_xlsxFiles()
    #startSAP()
    main()
