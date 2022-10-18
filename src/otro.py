import pandas as pd

def read_cuentas():
    cuentas=pd.read_excel("cuentas.xlsx",sheet_name="cuentas")
    return cuentas