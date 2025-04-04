import pandas as pd
import os



def extractData(filename):
    data = pd.read_excel(filename ,sheet_name = 'Resultados', engine = 'openpyxl', header = 3)

    data.iloc[:,4] = pd.to_datetime(data.iloc[:, 4])
    indexDate = 4
    indexClave = 9
    nameDate = data.columns[indexDate]
    nameClave = data.columns[indexClave]

    order = data.columns.to_list()
    lastest = data.sort_values(by = [nameDate, nameClave] ,ascending = [False, True])
    uniqueLastest = lastest.groupby(nameClave).first().reset_index()
    uniqueLastest = uniqueLastest[order]
    return uniqueLastest



def searchMatches(filename,parameters):
    template = pd.read_excel('template.xlsx', sheet_name ='FULL DATA', header = 2)
    data = extractData(filename)

    for index, item in data.iterrows():
        tipo = item.iloc[0]
        oD = item.iloc[1]
        wT = item.iloc[2]
        grado = item.iloc[3]
        proveedor = item.iloc[5]
        clave =  int(item.iloc[9])
        
        

    




searchMatches('SPC TPC TT04\SPC TPC L80\SPC TPC L80 1\SPC_TPC__93.17_x_12.20_L80_2017_TT04_-_PRUEBAS_(_OD_2_7-8_)_-__-_.xlsm', None)