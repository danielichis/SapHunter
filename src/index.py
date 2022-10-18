from processExtracts import process_xlsxFiles
from loadSap import startSAP, loadBankTemplates
from pathManagement import get_templates_path

def main():
    process_xlsxFiles() #procesamos todos los archivos en bruto y obtenemos sus plantillas
    sapInfo=get_templates_path() # con 4 ultimos digitos buscamos la info de cada plantilla para subir al sap
    startSAP() #iniciamos sap
    for template in sapInfo:
        loadBankTemplates(template) #cada template es un diccionario que tiene la ruta del archivo y la info de la cuenta

if __name__ == "__main__":
    main()