import requests 
import pandas as pd 
import numpy as np 
from pyxll import xl_macro, xl_app, xlcAlert

@xl_macro
def actualizar():

    # Excel Aprovisionamientos 
    xls = r'C:\Users\Usuario\Desktop\Solid2023\Excel Juan\REAPROVISIONAMIENTO_DEF.xlsm'

    # Dolibarr REAL
    api_key = 'clave_api123'
    url = "https://intranet.solidsoft-tray.com/api/index.php/" 

    # Dolibarr DEMO
    # api_key = 'nLBE1J8Tts3q99zy1B5VXl79SznAU6dl'
    # url = 'http://localhost/dolibarr/api/index.php/'

    # PREPARA LOS DATOS 
    concepto =  pd.read_excel(xls, header=1)['CONCEPTO'][0]
    if pd.isna(concepto): concepto = 'Sin concepto'
    dta = pd.read_excel(xls, header = 4).iloc[:,0:10]
    l = len(dta['Referencia'].dropna())
    dta = dta.iloc[0:l,:]
    base_url = f"{url}products?limit=500&DOLAPIKEY={api_key}"
    prods = requests.get(base_url).json() 

    # EXTRAE LAS REFERENCIAS Y LOS IDs
    ref_XLS = dta['Referencia'].dropna().values.tolist()    #Referencias en el excel
    ref_DB = [t['ref'] for t in prods]                      #Referencias en DB
    ids = [y['id'] for y in prods]                          #IDs en DB

    # COMPRUEBA LAS REFERENCIAS 
    FLAG = 0; not_in_DB = []
    for i in range(len(ref_XLS)):
        if ref_XLS[i] not in ref_DB: FLAG = 1; not_in_DB.append(ref_XLS[i])

    if FLAG == 1: xlcAlert(f'Algunas referencias del excel no están en dolibarr {not_in_DB}')

    # ACTUALIZA EL INVENTARIO EN DOLIBARR
    else: 

        # ASIGNA A CADA REF DEL EXCEL SU ID
        ids_XLS = []
        for k in range(len(ref_XLS)):
            for w in range(len(ref_DB)):
                if ref_DB[w] == ref_XLS[k]: ids_XLS.append(ids[w])

        # ACTUALIZA EL INVENTARIO
        for i in range(len(ref_XLS)):
                if ~np.isnan(dta['Albarán de Entrada'][i]) and dta['Albarán de Entrada'][i] != 0:

                    base_url = f"{url}stockmovements?DOLAPIKEY={api_key}"
                    res =  { "product_id": ids_XLS[i], "warehouse_id": 1, "qty": str(dta['Albarán de Entrada'][i]), "price": 0, "movementlabel": concepto} 
                    requests.post(base_url, json = res)


        # CALCULA EL INVENTARIO TOTAL ACTUALIZADO A PARTIR DE DB
        base_url = f"{url}products?limit=500&includestockdata=1&DOLAPIKEY={api_key}"
        stock_inicial = requests.get(base_url).json()
        stockv = [x["stock_theorique"] for x in stock_inicial]
        stock_inicial = [l["stock_warehouse"] for l in stock_inicial]
        stock = []; stockB = []; stockC = []; stockvA = []

        # Stock en almacen A
        for w in range(len(stock_inicial)):
            FLAG = 0
            if stock_inicial[w] == []: stock.append(0) # No hay stock en ningun almacen 
            else: 
                for i in range(len(list(stock_inicial[w]))): 
                    if list(stock_inicial[w])[i] == '1': stock.append(int(list(stock_inicial[w].values())[i]['real'])); FLAG = 1 #Hay stock en A
                if FLAG == 0: stock.append(0) # No hay stock en A
        
        # CALCULO DEL STOCK VIRTUAL DEL ALMACEN A:

        # Stock en almacen B 
        for w in range(len(stock_inicial)):
            FLAG = 0
            if stock_inicial[w] == []: stockB.append(0) # No hay stock en ningun almacen 
            else: 
                for i in range(len(list(stock_inicial[w]))): 
                    if list(stock_inicial[w])[i] == '2': stockB.append(int(list(stock_inicial[w].values())[i]['real'])); FLAG = 1 #Hay stock en B
                if FLAG == 0: stockB.append(0) # No hay stock en B

        # Stock en almacen C 
        for w in range(len(stock_inicial)):
            FLAG = 0
            if stock_inicial[w] == []: stockC.append(0) # No hay stock en ningun almacen 
            else: 
                for i in range(len(list(stock_inicial[w]))): 
                    if list(stock_inicial[w])[i] == '3': stockC.append(int(list(stock_inicial[w].values())[i]['real'])); FLAG = 1 #Hay stock en C
                if FLAG == 0: stockC.append(0) # No hay stock en C

        stockvA = [a - b - c for a, b, c in zip(stockv, stockB, stockC)]

        # STOCK DE ALERTA
        stock_alert = [y['seuil_stock_alerte'] for y in prods]

        # ASIGNA A CADA REF DEL EXCEL SU STOCK

        tstock = []; tstockv = []; astock = []
        for k in range(len(ref_XLS)):
            for t in range(len(ref_DB)):
                if ref_DB[t] == ref_XLS[k]: tstock.append(stock[t]); tstockv.append(stockvA[t]); astock.append(stock_alert[t])

        # ESCRIBE EL INVENTARIO TOTAL ACTUALIZADO EN LA EXCEL
        xl = xl_app()
        xlcAlert("Actualizando stock, por favor guarda el excel antes de realizar cambios")
        for k in range(len(ref_XLS)):

            cell = xl.Range(f'D{k+6}')                                                    # Escribe el stock en el excel
            if tstock[k] is None: cell.Value = 0 
            else: cell.Value = tstock[k]

            cell = xl.Range(f'F{k+6}')                                                    # Escribe el stock virtual en el excel
            if tstockv[k] is None: cell.Value = 0 
            else: cell.Value = tstockv[k]

            cell = xl.Range(f'G{k+6}')                                                    # Escribe el stock de alerta en el excel
            if astock[k] is None: cell.Value = 0
            else: cell.Value = astock[k]
        
            if ~np.isnan(dta['Albarán de Entrada'][k]):# and dta['Albarán de Entrada'][k] != 0:  # Pon a 0 el albarán de entrada 

                cell_stk = xl.Range(f'C{k+6}')
                cell_stk.Value = ''
            

        #xlcAlert("Stock actualizado con éxito, por favor guarda el excel antes de realizar cambios")