import pandas as pd
import os
import numpy as np


informacion = "info2.xlsx"
template = "template.xlsx"

def verify_Data():
    searchdata = pd.read_excel(template ,header=3, usecols="A:S")
    print("datos existentes")
    return searchdata


def parse_data():
    verify_Data()
    
    
    arrayinfo = pd.read_excel(informacion, header=3, sheet_name="Resultados", usecols="A,B,C,E,D,J,F,X,Y,Z,AB,AC,AD,AE,AF,AG,AH")

    print("libro leido con exito") 
    
    verArray = arrayinfo.style.to_html()

    html_content = f"""<!DOCTYPE html>
    <html>
    <head>
        <title>My DataFrame</title>
    </head>
    <body>F
        {verArray}
    </body>
    </html>"""

    clavesUnicas = arrayinfo['Clave'].unique()
    print(clavesUnicas)
    
    with open('table.html', 'w') as f:
        f.write(html_content)


    return arrayinfo


    
    
parse_data()