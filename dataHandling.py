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



def searchMatches(filename):
    template = pd.read_excel('template.xlsx', sheet_name ='FULL DATA', header = 2)
    data = extractData(filename)

    for index, item in data.iterrows():
        tipo = item.iloc[0]
        oD = item.iloc[1]
        wT = item.iloc[2]
        grado = item.iloc[3]
        proveedor = item.iloc[5]
        clave =  int(item.iloc[9])
        item = pd.DataFrame(item)
        
        ifod = template['OD'] == oD
        ifwt = template['WT'] == wT
        ifgrado = template['Grado'] == grado
        ifclave = template['Clave'] == clave

        ifAll = (ifod & ifwt & ifgrado & ifclave)

        matches = template[ifAll]
        index = matches.index[0]
        #matches.loc[index] = item
        nuevaData = item
        vCalentamientol1Austenizado = nuevaData.loc['L1 °C']
        vCalentamientol2Austenizado = nuevaData.loc['L2 °C']
        vEmpapeL1Austenizado = nuevaData.loc['L1 °C.1']
        vVelPasaje = nuevaData.loc['m/s']
        vFlujo = nuevaData.loc['m3/h']
        vCalentamientol1Revenido = nuevaData.loc['L1 °C.2']
        vCalentamientol2Revenido = nuevaData.loc['L2 °C.1']
        vEmpapeL1Revenido = nuevaData.loc['L1 °C.3']
        vEmpapeL2Revenido = nuevaData.loc['L2 °c']
        vTCs = nuevaData.loc['seg.1']
        #print(index)

        template.loc[index, ['Calentamiento L1 austenizado','Calentamiento L2 austenizado','Empape L2 austenizado','Vel. Pasaje','Flujo','Calentamiento L1 revenido','Calentamiento L2 revenido','Empape L1 revenido','Empape L2 revenido','TC [s]']] = [vCalentamientol1Austenizado,vCalentamientol2Austenizado,vEmpapeL1Austenizado,vVelPasaje,vFlujo,vCalentamientol1Revenido,vCalentamientol2Revenido,vEmpapeL1Revenido,vEmpapeL2Revenido,vTCs] 
        print(item)
        print('-'*30)
        print(f'datos copiados con exito en la fila # {index}')

    template.to_excel('templateNew.xlsx', header = True, index = False)



searchMatches('SPC TPC TT04\SPC TPC L80\SPC TPC L80 1\SPC_TPC__93.17_x_12.20_L80_2017_TT04_-_PRUEBAS_(_OD_2_7-8_)_-__-_.xlsm')