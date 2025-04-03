import pandas as pd
import numpy as np
import logging
import os

logging.basicConfig(filename='registro.log',  
                    level=logging.INFO,    
                    format='%(asctime)s - %(levelname)s - %(message)s')

template = "template.xlsx"

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
    return uniqueMostRecentValues


def convertMmtoIn(x):
    y = x / 25.4
    y = round(y, 3)
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
    datos = parse_data(informacion)
    datosformateados = datos.to_numpy()
    #print(datosformateados)
    
    lookUpValues = pd.read_excel(template, header = 0, engine='openpyxl') 
    
    try:
        indexArray = 0
        for d in datosformateados:
                oD = convertMmtoIn(datosformateados[indexArray,2])
                wT = convertMmtoInWT(datosformateados[indexArray,3])
                tipo = convTipo(datosformateados[indexArray, 1])
                clave = datosformateados[indexArray,0]
                grado = None
                if datosformateados[indexArray, 4] == 'N80Q':
                            grado = 'N80'
                elif datosformateados[indexArray, 4] == 'P110 ICY' or datosformateados[indexArray, 4] == 'P110 IC' or datosformateados[indexArray, 4] == 'P110 ICCY':
                            grado = 'P110 ICCY'
                elif datosformateados[indexArray, 4] == 'L80CR1':
                            grado = 'L80 CR1'
                elif datosformateados[indexArray, 4] == 'L80ICY':
                            grado = 'L80 ICY'
                elif datosformateados[indexArray, 4] == 'TN110HC':
                            grado = 'TN110 HC'
                else: grado = datosformateados[indexArray, 4]

                unmatched_data = {
                        "Clave": clave,
                        "OD": oD,
                        "WT": wT,
                        "Tipo": tipo,
                        "Grado": grado
                    }


                condicion = (lookUpValues['WT'] >=limiteAbajo(wT, 0.02) )  & (lookUpValues['WT'] <=limiteArriba(wT, 0.02) ) &(lookUpValues['OD'] >=limiteAbajo(oD, 0.5) ) &(lookUpValues['OD'] <=limiteArriba(oD, 0.5) )  & (lookUpValues['TIPO'] == tipo) & (lookUpValues['Clave'] == clave) &  (lookUpValues['Grado'] == grado) 
                itemunico = lookUpValues[condicion]

                if len(itemunico) > 1:
                    print(f"Se encontraron muchas coincidencias en la template con estos datos {clave}, {oD}, {wT}, {grado}")
                    print("deseas omitirlos 1. Si o 2. No")     
                    while True:
                        opcion = input("Elige una opcion (1 o 2): ")   
                        if opcion == '1':
                                    print("omitiste este registro")
                                    break  
                        elif opcion == '2':
                                print("Saliendo del procesamiento.")
                                return "por favor edita manualmente el archivo"
                        else:
                                print("Opción inválida. Por favor, introduce 1 o 2.")
                #print(d)
                if itemunico.empty:
                    items = []
        
                    faltan = 0

                    condicion_wt = (lookUpValues['WT'] >=limiteAbajo(wT, 0.05) )  &  (lookUpValues['WT'] <=limiteArriba(wT, 0.05) )
                    if not any(condicion_wt):
                       items.append('WT') 
                       faltan = faltan + 1

                    condicion_od = (lookUpValues['OD'] >=limiteAbajo(oD, 0.5) ) & (lookUpValues['OD'] <=limiteArriba(oD, 0.5) )
                    if not any(condicion_od):
                        items.append('OD') 
                        faltan = faltan + 1

                    condicion_tipo = (lookUpValues['TIPO'] == tipo)
                    if not any(condicion_tipo):
                        items.append('Tipo') 
                        faltan = faltan + 1

                    condicion_clave = (lookUpValues['Clave'] == clave)
                    if not any(condicion_clave):
                        items.append('Clave') 
                        faltan = faltan + 1

                    condicion_grado = (lookUpValues['Grado'] == grado)
                    if not any(condicion_grado):
                        items.append('Grado') 
                        faltan = faltan + 1

                    print(f"Faltan {faltan} parámetros para una coincidencia del 100%.")
                    print(items)


                    #print(d)
                    #print(unmatched_data, "\n")
                    print("No se encontro una coincidencia de los datos en la template")
                    print("deseas omitirlos 1. Si o 2. No")
                    while True:
                        opcion = input("Elige una opcion (1 o 2): ")
                        if opcion == '1':
                            nombre_archivo = os.path.basename(informacion)
                            rutaArchivo = os.path.dirname(informacion)
                            registro_omitido = f"Clave: {clave}, OD: {oD}, WT: {wT}, Tipo: {tipo}, Grado: {grado}"
                            logging.info(f"Archivo: {nombre_archivo} - Registro omitido (sin coincidencia): {registro_omitido}")
                            a = open('registrossincoincidencias.txt', 'a')
                            a.write(f"-Values:  {registro_omitido} ---- {rutaArchivo} \n")
                            a.close()
                            print("omitiste este registro")
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
    except Exception as e:
        raise e
    
    
def procesar_carpeta(ruta):
    
    try:
        index = 0
        for (root, subcarpetas, archivos) in os.walk(ruta):
            for archivo in archivos:
                if archivo.endswith((".xlsx", ".xls", ".xlsm")):
                    ruta_completa = os.path.join(root, archivo)
                    print(ruta_completa)
                    index = index + 1
                    try:
                        verify_Data(ruta_completa)
                    except Exception as e:
                        print(f"Se encontró un error al procesar el archivo: {archivo}")
                        print(f"Razón del error: {e}")
                        print("Quieres omitir este archivo y continuar con el siguiente? 1. Sí o 2. No")
                        while True:
                            opcion = input("Elige una opcion (1 o 2): ")
                            if opcion == '1':
                                print("Omitiendo archivo.")
                                logging.info(f"Archivo: {archivo} - Omitido debido a error: {e}")
                                a = open('archivosconerrores.txt', 'a')
                                a.write( f'- {ruta_completa} \n')
                                a.close()
                                break  
                            elif opcion == '2':
                                print("Has seleccionado no omitir.")
                                return f"Se encontraron errores. Por favor, revisa el archivo: {archivo} - Razón: {e}"
                            else:
                                print("Opcion inválida")

        print(f"\nProceso completado para todos los archivos Excel en la carpeta: {ruta}, se procesaron {index} archivos")
        w = open('carpetasrevisadas.txt','a')
        w.write(f'- {ruta} - ({index})\n')
        w.close()



    except FileNotFoundError:
        print(f"Error: No se encontró la carpeta {ruta}")
        return f"Error: No se encontró la carpeta {ruta}"


    
    
#procesar_carpeta("S:/SEC/TTRTUCA/0- TT04 Master Plan/SPC TT04 2015/SPC CASING TT04")


def buscar_con_parametros_en_arbol_excel(ruta_raiz, parametros_busqueda):

    all_matches = []
    try:
        for dirpath, dirnames, filenames in os.walk(ruta_raiz):
            for filename in filenames:
                if filename.endswith(('.xlsx', '.xls', 'xlsm')):
                    ruta_archivo = os.path.join(dirpath, filename)
                    try:
                        df = pd.read_excel(ruta_archivo,header=3, sheet_name="Resultados", engine = 'openpyxl')
                        df['Fecha'] = pd.to_datetime(df['Fecha'])
                        condiciones = pd.Series([True] * len(df)) 

                        for columna, texto_busqueda in parametros_busqueda.items():
                            if columna == 'Clave':
                                    if 'Clave' in df.columns:
                                        columna = 'Clave'
                                    elif 'CLAVE' in df.columns:
                                        columna = 'CLAVE'
                                    elif 'clave' in df.columns:
                                        columna = 'clave'
                            if columna in df.columns:
                                df[columna] = df[columna].astype(str)
                                condicion_columna = df[columna].str.contains(texto_busqueda, case=False, na=False)
                                condiciones = condiciones & condicion_columna 

                            else:
                                print(f"Advertencia: La columna '{columna}' no existe en el archivo '{ruta_archivo}'.")
                                condiciones = pd.Series([False] * len(df)) 

                        resultados = df[condiciones]

                        if not resultados.empty:
                            resultados['Archivo_Origen'] = os.path.relpath(ruta_archivo, ruta_raiz)
                            all_matches.append(resultados)

                    except Exception as e:
                        print(f"Error al leer el archivo '{ruta_archivo}': {e}")

        if all_matches:
            df_resultados_final = pd.concat(all_matches, ignore_index=True)
            ruta_archivo_salida = os.path.join('resultados.xlsx') 
            arrayordenadoDate = df_resultados_final.sort_values( by='Fecha', ascending=False)
            arrayordenadoDate.to_excel(ruta_archivo_salida, index=False)
            print(f"\nSe encontraron coincidencias. Los resultados se han guardado en: {ruta_archivo_salida}")
            asd = arrayordenadoDate.iloc[0]
            iop = asd['Archivo_Origen']
            qwe = iop.replace("\\", "/")
            return qwe

        else:
            print("No se encontraron coincidencias con los parámetros especificados.")
            return None

    except FileNotFoundError:
        print(f"Error: La carpeta raíz '{ruta_raiz}' no fue encontrada.")
        return pd.DataFrame()
    except Exception as e:
        print(f"Ocurrió un error al procesar la carpeta y sus subdirectorios: {e}")
        return pd.DataFrame()


ruta_de_la_carpeta_raiz = 'S:/SEC/TTRTUCA/0- TT04 Master Plan/SPC TT04 2015/SPC CASING TT04' 



parametros_de_busqueda = {
    'Grado': 'N80',  
    'Clave': '302',
    'OD (mm)':'177',
    'WT (mm)': '11',
    'Proveedor': 'TAMSA' 
}


buscar_con_parametros_en_arbol_excel(ruta_de_la_carpeta_raiz, parametros_de_busqueda)

#print(buscar_con_parametros_en_arbol_excel(ruta_de_la_carpeta_raiz, parametros_de_busqueda))
#verify_Data(f'S:/SEC/TTRTUCA/0- TT04 Master Plan/SPC TT04 2015/SPC CASING TT04/{buscar_con_parametros_en_arbol_excel(ruta_de_la_carpeta_raiz, parametros_de_busqueda)}')



