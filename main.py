import pandas as pd
import os
import numpy as np



informacion = "info2.xlsx"
template = "template.xlsx"



def parse_data():
    
    arrayinfo = pd.read_excel(informacion, header=3, sheet_name="Resultados", usecols="A,B,C,E,D,J,F,X,Y,Z,AB,AC,AD,AE,AF,AG,AH")
    
    arrayinfo['Fecha'] = pd.to_datetime(arrayinfo['Fecha'])
    arrayordenadoDate = arrayinfo.sort_values(by = ['Fecha', 'Clave'], ascending = [False, True])


    
    


    
    uniqueMostRecentValues = arrayordenadoDate.groupby('Clave').first().reset_index()

    verArray = uniqueMostRecentValues.style.to_html()
    
    
    html_content = f"""<!DOCTYPE html>
    <html>
    <head>
        <title>My DataFrame</title>
    </head>
    <body>F
        {verArray}
    </body>
    </html>"""
    
    
    with open('table.html', 'w') as f:
        f.write(html_content)

    return uniqueMostRecentValues


def convertMmtoIn(x):
    y = x / 25.4
    return y


def convertMmtoInWT(x):
    x = x / 25.4
    y = round(x, 3)
    return y


def convTipo(tipo):
    tipo = tipo[0:3]
    tipo = tipo.upper()
    return tipo





def verify_Data():
    datos = parse_data()
    datosformateados = datos.to_numpy()
    #print(datos)

    
    lookUpValues = pd.read_excel(template, header = 2) 
    
    #print(lookUpValues)
    
    indexArray = 0
    for d in datosformateados:
        oD =  convertMmtoIn(datosformateados[indexArray,2])
        wT = convertMmtoInWT(datosformateados[indexArray, 3])
        tipo = convTipo(datosformateados[indexArray, 1])
        clave = datosformateados[indexArray,0]
        grado = datosformateados[indexArray, 4]
        
        condicion = (lookUpValues['WT'] == wT) & (lookUpValues['OD'] == oD) & (lookUpValues['TIPO'] == tipo) & (lookUpValues['Clave'] == clave) & (lookUpValues['Grado'] == grado)
        itemunico = lookUpValues[condicion]
        indicefila = itemunico.index[0]
        
        itemunico.loc[indicefila]
        
        
        
        print(itemunico)
        
        print(d)
        
        
        
        
        indexArray = indexArray + 1
        #print(indexArray, oD, wT, clave, tipo)
    
    
    
    searchdata = pd.read_excel(template ,header=3, usecols="A:S")
    
    
    
    
    return searchdata


    
verify_Data()