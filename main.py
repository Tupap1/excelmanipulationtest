import pandas as pd
import os
import numpy as np
import logging
import webbrowser
#import onedrivesdk
import msal
from dotenv import load_dotenv


def obtener_Token_acces(aplicationId, clientSecret, scopes):
    client = msal.ConfidentialClientApplication(
        client_id = aplicationId,
        client_credential = clientSecret,
        authority = 'https://login.microsoftonline.com/consumers/' 
        
    )
    
    auth_request_url= client.get_authorization_request_url(scopes)
    webbrowser.open(auth_request_url)
    autorizationCode = input("enter ur autorization code: ") 
    
    tokenResponse = client.acquire_token_by_authorization_code(
        code = autorizationCode,
        scopes = scopes
    )
    
    if 'access_code' in tokenResponse:
        return tokenResponse['access_token']
    else:
        return ' no se pudo obtener el token ' + str(tokenResponse)

    





informacion = "info2.xlsx"
template = "template.xlsx"
logging.basicConfig(filename='registro.log',  
                    level=logging.INFO,    
                    format='%(asctime)s - %(levelname)s - %(message)s')


def parse_data():
    
    arrayinfo = pd.read_excel(informacion, header=3, sheet_name="Resultados", usecols="A,B,C,E,D,J,F,X,Y,Z,AB,AC,AD,AE,AF,AG,AH")
    
    arrayinfo['Fecha'] = pd.to_datetime(arrayinfo['Fecha'])
    arrayordenadoDate = arrayinfo.sort_values(by = ['Fecha', 'Clave'], ascending = [False, True])

    uniqueMostRecentValues = arrayordenadoDate.groupby('Clave').first().reset_index()

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

    
    lookUpValues = pd.read_excel(template, header = 0) 
    
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
        
        
        nuevaData = datos.iloc[0]
        vCalentamientol1 = nuevaData.iloc[7]
        vCalentamientol2 = nuevaData.iloc[8]
        vEmpapeL2 = nuevaData.iloc[9]
        vVelPasaje = nuevaData.iloc[10]
        vFlujo = nuevaData.iloc[11]
        vCalentamientol1_1 = nuevaData.iloc[12]
        vCalentamientol2_1 = nuevaData.iloc[13]
        vEmpapeL1 = nuevaData.iloc[14]        
        vEmpapeL2_1 = nuevaData.iloc[15]
        vTCs = nuevaData.iloc[16]
        
        lookUpValues.loc[indicefila, ['Calentamiento L1','Calentamiento L2','Empape L2','Vel. Pasaje','Flujo','Calentamiento L1.1','Calentamiento L2.1','Empape L1','Empape L2.1','TC [s]']] = [vCalentamientol1,vCalentamientol2,vEmpapeL2,vVelPasaje,vFlujo,vCalentamientol1_1,vCalentamientol2_1,vEmpapeL1,vEmpapeL2_1,vTCs]
        
        asd = str(lookUpValues.iloc[indicefila])

        lookUpValues.to_excel(template, index=False, header=True)
        
        logging.info(f"registro agregado:  {asd}" )
        print("operacion realizada con exito", indicefila)
        indexArray = indexArray + 1
    
    
    
    
    
    
    

def main():
    load_dotenv()
    APPLICATION_ID =  os.getenv('APLICATION_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    SCOPES = ['User.Read', 'User.Read']
    
    try:
        access_token = obtener_Token_acces(aplicationId=APPLICATION_ID, clientSecret=CLIENT_SECRET,scopes= SCOPES)
        headers = {
            'autorization': 'Bearer: ' + access_token
        }
        print(headers)
    except Exception as e:
        print(f'error: {e}')
        
    
    
    #verify_Data()
    
main()