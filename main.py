import pandas as pd
import numpy as np
import logging
import os





#informacion = "info2.xlsx"
template = "template.xlsx"
logging.basicConfig(filename='registro.log',  
                    level=logging.INFO,    
                    format='%(asctime)s - %(levelname)s - %(message)s')


def parse_data(informacion):


    arrayinfo = pd.read_excel(informacion, header=3, sheet_name="Resultados", usecols="A,B,C,E,D,J,F,X,Y,Z,AB,AC,AD,AE,AF,AG,AH", engine='openpyxl')
    
    arrayinfo['Fecha'] = pd.to_datetime(arrayinfo['Fecha'])

    columna_clave = None
    if 'Clave' in arrayinfo.columns:
        columna_clave = 'Clave'
    elif 'CLAVE' in arrayinfo.columns:
        columna_clave = 'CLAVE'
    elif 'clave' in arrayinfo.columns:
        columna_clave = 'clave'

    arrayordenadoDate = arrayinfo.sort_values(by = ['Fecha', columna_clave], ascending = [False, True])

    uniqueMostRecentValues = arrayordenadoDate.groupby(columna_clave).first().reset_index()

    print(uniqueMostRecentValues)
    print('datos a copiar')
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



def limiteAbajo( x, tolerancia):
    lim =  x - tolerancia
    return lim

def limiteArriba( x,tolerancia):
    lim = x + tolerancia 
    return lim



def verify_Data(informacion):
    unmatched_records_in_file = []
    datos = parse_data(informacion)
    datosformateados = datos.to_numpy()

    
    lookUpValues = pd.read_excel(template, header = 0, engine='openpyxl') 
    
    #print(lookUpValues)
    try:
        indexArray = 0
        for d in datosformateados:
                oD =  convertMmtoIn(datosformateados[indexArray,2])
                wT = convertMmtoInWT(datosformateados[indexArray, 3])
                tipo = convTipo(datosformateados[indexArray, 1])
                clave = datosformateados[indexArray,0]
                grado = datosformateados[indexArray, 4]

                condicion = (lookUpValues['WT'] >=limiteAbajo(wT, 0.05) )  & (lookUpValues['WT'] <=limiteArriba(wT, 0.05) ) &(lookUpValues['OD'] >=limiteAbajo(oD, 0.5) ) &(lookUpValues['OD'] <=limiteArriba(oD, 0.5) )  & (lookUpValues['TIPO'] == tipo) & (lookUpValues['Clave'] == clave)
                itemunico = lookUpValues[condicion]
                print(itemunico)
                print("obteniendo fila")
                if itemunico.empty:
                    unmatched_data = {
                        "Clave": clave,
                        "OD": oD,
                        "WT": wT,
                        "Tipo": tipo,
                        "Grado": grado
                    }
                    print(d)
                    print(unmatched_data)
                    print("No se encontro una coincidencia de los datos en la template")
                    print("deseas omitirlos 1. Si o 2. No")
                    while True:
                        opcion = input("Elige una opcion (1 o 2): ")
                        if opcion == '1':
                            print("omitiste este registro")
                            nombre_archivo = os.path.basename(informacion)
                            registro_omitido = f"Clave: {clave}, OD: {oD}, WT: {wT}, Tipo: {tipo}, Grado: {grado}"
                            logging.info(f"Archivo: {nombre_archivo} - Registro omitido (sin coincidencia): {registro_omitido}")
                            a = open('registrossincoincidencias.txt\n', 'a')
                            a.write("- " ,registro_omitido, "\n")
                            a.close()
                            indexArray = indexArray + 1  
                            break  
                        elif opcion == '2':
                            print("Saliendo del procesamiento.")
                            return "por favor edita manualmente el archivo"
                        else:
                            print("Opción inválida. Por favor, introduce 1 o 2.")
                else:
                    indicefila = itemunico.index[0]
                    print('donde copiarlos', indicefila)

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
                
                return unmatched_records_in_file
        
    except Exception as e:
        raise e
    
    
def procesar_carpeta_excel(ruta_carpeta):

    try:
        archivos_en_carpeta = os.listdir(ruta_carpeta)
        for nombre_archivo in archivos_en_carpeta:
            if nombre_archivo.endswith((".xlsx", ".xls", ".xlsm")):
                ruta_completa = os.path.join(ruta_carpeta, nombre_archivo)
                print(nombre_archivo)
                try:
                    verify_Data(ruta_completa)
                except Exception as e:
                    print(f"Se encontró un error al procesar el archivo: {nombre_archivo}")
                    print(f"Razón del error: {e}")
                    print("Quieres omitir este archivo y continuar con el siguiente? 1. Sí o 2. No")
                    while True:
                        opcion = input("Elige una opcion (1 o 2): ")
                        if opcion == '1':
                            print("Omitiendo archivo.")
                            logging.info(f"Archivo: {nombre_archivo} - Omitido debido a error: {e}")
                            break  
                        elif opcion == '2':
                            print("Has seleccionado no omitir. ")
                            return f"Se encontraron errores. Por favor, revisa el archivo: {nombre_archivo} - Razón: {e}"
                        else:
                            print("Opción inválida. Por favor, introduce 1 o 2.")
        print(f"\nProceso completado para todos los archivos Excel en la carpeta: {ruta_carpeta}")



    except FileNotFoundError:
        print(f"Error: No se encontró la carpeta {ruta_carpeta}")
        return f"Error: No se encontró la carpeta {ruta_carpeta}"

    
    
procesar_carpeta_excel("S:/SEC/TTRTUCA/0- TT04 Master Plan/SPC TT04 2015/SPC TUBING TT04/SPC N80Q Tubing TT04")
#verify_Data("S:/SEC/TTRTUCA/0- TT04 Master Plan/SPC TT04 2015/SPC CASING TT04/SPC L80 Casing TT04/SPC CL8070362 SMLS  2021 TT04 - Pruebas - .xlsm")