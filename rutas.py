import os

def xlsmToxlsx(ruta_archivo):

    if ruta_archivo.endswith(".xlsm"):
        ruta_xlsx = ruta_archivo[:-5] + ".xlsx"
        try:
            os.rename(ruta_archivo, ruta_xlsx)
            print(f"Archivo renombrado: '{ruta_archivo}' a '{ruta_xlsx}'")
            return ruta_xlsx
        except FileNotFoundError:
            print(f"Error: Archivo no encontrado en la ruta '{ruta_archivo}'")
            return None
        except OSError as e:
            print(f"Error al renombrar '{ruta_archivo}': {e}")
            return None
    else:
        print(f"Advertencia: El archivo '{ruta_archivo}' no tiene la extensión .xlsm (sensible a mayúsculas/minúsculas).")
        return None
    


def xlsxToxslm(ruta_archivo):
    if ruta_archivo.endswith(".xlsx"):
        ruta_xlsm = ruta_archivo[:-5] + ".xlsm"
        try:
            os.rename(ruta_archivo, ruta_xlsm)
            print(f"Archivo renombrado: '{ruta_archivo}' a '{ruta_xlsm}'")
            return ruta_xlsm
        except FileNotFoundError:
            print(f"Error: Archivo no encontrado en la ruta '{ruta_archivo}'")
            return None
        except OSError as e:
            print(f"Error al renombrar '{ruta_archivo}': {e}")
            return None
    else:
        print(f"Advertencia: El archivo '{ruta_archivo}' no tiene la extensión .xlsx (sensible a mayúsculas/minúsculas).")
        return None