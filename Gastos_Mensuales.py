# Cómo importar las librerías
import xlwings  as xw
import numpy    as np
import pandas   as pd
import datetime as dt
import os
import sys
from pathlib import Path
import re

# Gráficos
# ==============================================================================
import matplotlib.pyplot as plt
import seaborn as sns
#%matplotlib inline
plt.style.use('fivethirtyeight')

#from statsmodels.graphics.tsaplots import plot_acf
#from statsmodels.graphics.tsaplots import plot_pacf


# Configuración warnings
# ==============================================================================
import warnings
warnings.filterwarnings('ignore')

# Configuracion Pandas
pd.set_option('display.max_rows', 200)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', -1)
pd.set_option('display.float_format', '{:.2f}'.format)


@xw.func
def hello(name):
    return f"Hello {name}!"

def main():
    # GESTION GASTOS  Fecha: 01/03/2022 07:00
    '''
    # -*- coding: utf-8 -*-
    Created on Martes 8 Marzo 2022 

    @author: Daniel Busso 
    '''
    # GENERAMOS PLANILLA MOVIMIENTOS DEL MES (POR AHORA SE LLAMA Febrero)
    #  
    # EMPIEZO PROCESANDO UNA CARPETA
    # 
    # ELEGIR PLATAFORMA WINDOWS LOCAL O GOOGLE COLABORATORY
    if sys.platform == 'linux' :
        from google.colab import drive
        drive.mount('/content/drive', force_remount=True)
        carpeta_trabajo = Path('/content/drive/Mi unidad/Mis_Pagos/Planilla_Gastos') # Google Colab
    else:
        #carpeta_inicio  = Path.home()
        #carpeta_gg = Path("C:/users/debus/GestionGastos")
        #carpeta_trabajo = carpeta_inicio / carpeta_gg
        libro = xw.Book.caller()
        periodo = libro.sheets('Menu').range('H2').value
        mes  = periodo[0:2]
        anio = periodo[2:]
        carpeta = Path(r'G:\Mi unidad\Mis_Pagos', anio, periodo)
        Archivo = list(carpeta.glob('CA*$*0020382809*.xls')) # OJO DEBE HABER SOLO UN ARCHIVO POR CARPETA

    # GUARDAR EL ARCHIVO CA $ 900 0020382809-Movimientos.xls COMO Movi.xlsx
    # Archivo = carpeta_trabajo + '/' + 'Movi.xls'  HAY QUE VERLO

    # NO PUEDO USAR decimal=',', thousands='.' debido al formato que tiene la columna "Monto" Por ejemplo $1.500,45
    # datos = pd.read_excel(Archivo, names=['Fecha', 'Concepto', 'Descripcion', 'Monto', 'NroComprobante'], 
                        # skiprows=5, skipfooter=2, decimal=',', thousands='.', engine='openpyxl')
    
    # =============================================================================================
    # LEER ARCHIVO DE MOVIMIENTOS CA PESOS BANCOR y GENERAMOS EL DATAFRAME DEL MES 
    #
    datos = pd.read_excel(Archivo[0], names=['Fecha', 'Concepto', 'Descripcion', 'Monto', 'NroComprobante'], 
                        skiprows=5, skipfooter=2, engine='xlrd')  
    datos['Fecha'] = pd.to_datetime(datos['Fecha'], format='%d/%m/%Y')

    datos['Monto'] = datos['Monto'].str.lstrip('$')
    datos['Monto'] = datos['Monto'].str.strip()
    datos['Monto'] = datos['Monto'].str.replace('.', '',  regex=False)  # Elimina los puntos de miles
    datos['Monto'] = datos['Monto'].str.replace(',', '.', regex=False)  # Reemplaza la coma decimal por el punto decimal
    datos['Monto'] = pd.to_numeric(datos['Monto'])  # Pasa de type Object a Float
    #datos['Monto'] = datos['Monto'].astype(float)

    datos['NroComprobante'] = datos['NroComprobante'].astype(str)

    archivo_cp = Path(r'G:\Mi unidad\Mis_Pagos', anio, periodo, 'Movi.xlsx')
    datos.to_excel(archivo_cp, engine='openpyxl')
    Movi = datos.copy()

    datos_ord = Movi.sort_values(by='Concepto', ascending=True)

    #Mes = Movi.loc[0, 'Fecha'].strftime('%B')  # Febrero

    # ============================================================================================= 
    # ESCRIBIMOS LOS MOVIMIENTOS BANCOR A LA PLANILLA "MES" 
    #planilla = xw.Book.caller().sheets[0]

    periodo = libro.sheets('Menu').range('H2').value
    mes  = periodo[0:2]
    anio = periodo[2:]
    Mes_dic = {'01':'Ene', '02':'Feb', '03':'Mar', '04':'Abr', '05':'May', '06':'Jun', '07':'Jul', '08':'Ago', '09':'Sep', '10':'Oct', '11':'Nov', '12':'Dic' }
    Mes = Mes_dic[mes]

    Mes_lis = ['Falso', 'Ene', 'Feb', 'Mar','Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
    if Mes == 'Ene' :
        planilla = 'Menu'
    else:
        Mes_ant = Mes_lis[(int(mes) - 1)]
        planilla = Mes_ant + '_Efectivo'

    libro.sheets('_xlwings.conf').visible=0  # 0 OCULTA
    #libro = xw.Book()
    libro.sheets.add(name=Mes, after=libro.sheets(planilla)) # OJO TODAS LAS PLANILLAS SE VAN AGREGAR DESPUES DE ENERO
    libro.sheets(Mes).range('A1').options(pd.DataFrame, numbers='0,00', dates=dt.date, expand='table').value = datos_ord
    table = libro.sheets(Mes).tables.add(source=libro.sheets(Mes).range('A1').expand(), name=Mes)
    
    libro.sheets(Mes).autofit()
    libro.sheets(Mes).range('A:A').column_width=3   # indice
    libro.sheets(Mes).range('B:B').column_width=10  # Fecha
    libro.sheets(Mes).range('C:C').column_width=30  # Concepto
    libro.sheets(Mes).range('D:D').column_width=40  # Descripcion
    libro.sheets(Mes).range('E:E').column_width=12  # Monto
    libro.sheets(Mes).range('F:F').column_width=10  # NroComprobante
    libro.sheets(Mes).range('E:E').number_format = '0,00'
    

    # =============================================================================================
    # GENERAMOS PLANILLA DEBITOS
    # ====> Genero el DataFrame
    d = Movi.loc[Movi["Monto"] <= 0, :]
    debitos = d.copy()

    debitos['Tarjetas'] = np.nan

    # UNA FORMA: ES cREAR PRIMERO LA COLUMNA Y DESPEJAR EL VALOR Y SU INDICE Y GUARDARLO USANDO LOC
    #tarjeta = debitos.loc[debitos['Descripcion'].str.contains('TARJNARANJA.+', regex=True, na=False), 'Monto'] # QUE PASA SI TRAE MAS DE 1 REGISTRO
    #debitos.loc[tarjeta.index, 'Tarjetas'] = tarjeta
    #debitos.loc[tarjeta.index, 'Monto']    = np.nan

    # OTRA FORMA MAS SEGURA
    # ES SEPARAR CADA CONCEPTO EN UNA COLUMNA UNICA Y LUEGO SUMAR O COMBINAR ESAS COLUMNAS EN UNA SOLA
    # NO DA ERROR SI LA BUSQUEDA TRAE MAS DE UN REGISTRO

    debitos['Tarjetas'] = debitos.loc[debitos['Descripcion'].str.contains('TARJNARANJA.+', regex=True, na=False) | \
                                      debitos['Descripcion'].str.contains('VISA.+', regex=True, na=False), 'Monto']

    debitos.loc[debitos['Descripcion'].str.contains('TARJNARANJA.+', regex=True, na=False), 'Monto'] = np.nan
    debitos.loc[debitos['Descripcion'].str.contains('VISA.+', regex=True, na=False), 'Monto'] = np.nan 

    debitos['Impuestos'] = np.nan

    debitos['Inversiones'] = np.nan
    debitos['Inversiones'] = debitos.loc[debitos['Concepto'].str.contains('D.bito.+plazo fijo.+', regex=True, na=False) | \
                                         debitos['Concepto'].str.contains('D.bito operaci.n.+cambio ME.+', regex=True, na=False) | \
                                         debitos['Concepto'].str.contains('IMPUESTO P.A.I.S.+', regex=True, na=False) | \
                                         debitos['Concepto'].str.contains('.+GAN.+TENENCIA.+4815/20.+', regex=True, na=False), 'Monto']
    
    debitos.loc[debitos['Concepto'].str.contains('D.bito.+plazo fijo.+', regex=True, na=False), 'Monto'] = np.nan
    debitos.loc[debitos['Concepto'].str.contains('D.bito operaci.n.+cambio ME.+', regex=True, na=False), 'Monto']  = np.nan
    debitos.loc[debitos['Concepto'].str.contains('IMPUESTO P.A.I.S.+', regex=True, na=False), 'Monto'] = np.nan
    debitos.loc[debitos['Concepto'].str.contains('.+GAN.+TENENCIA.+4815/20.+', regex=True, na=False), 'Monto'] = np.nan

    debitos['Extracciones'] = np.nan
    debitos['Extracciones'] = debitos.loc[debitos['Concepto'].str.contains('Extracci.n.+', regex=True, na=False), 'Monto']
    debitos.loc[debitos['Concepto'].str.contains('Extracci.n.+', regex=True, na=False), 'Monto'] = np.nan

    del debitos['NroComprobante']

    debitos['Comentarios'] = ''

    debitos = debitos.rename(columns={'Monto':'Gastos'})

    debitos_ord = debitos.sort_values(by=['Gastos', 'Tarjetas', 'Inversiones', 'Extracciones'], ascending=False)

    #v_suma_gastos_deb       = debitos['Gastos'].sum()
    #v_suma_tarjetas_deb     = debitos['Tarjetas'].sum()
    #v_suma_impuestos_deb    = debitos['Impuestos'].sum()
    #v_suma_inversiones_deb  = debitos['Inversiones'].sum()
    #v_suma_extracciones_deb = debitos['Extracciones'].sum()
    # debitos.loc[debitos.index.max()+1, :] = [np.nan, np.nan, np.nan, v_suma_gastos, v_suma_tarjetas, v_suma_impuestos, v_suma_inversiones, v_suma_extracciones, '']

    # =====> Genero Planilla
    #mes = Movi.loc[0, 'Fecha'].strftime('%b')  # Feb 
    mes_deb = Mes + '_Debitos'
    libro.sheets.add(name=mes_deb, after=libro.sheets(Mes))
    libro.sheets(mes_deb).range('A1').options(pd.DataFrame, numbers='0,00', dates=dt.date, expand='table').value = debitos_ord
    table = libro.sheets(mes_deb).tables.add(source=libro.sheets(mes_deb).range('A1').expand(), name=mes_deb)
    libro.sheets(mes_deb).autofit()
    libro.sheets(mes_deb).range('A:A').column_width=3
    libro.sheets(mes_deb).range('B:B').column_width=10
    libro.sheets(mes_deb).range('C:C').column_width=30
    libro.sheets(mes_deb).range('D:D').column_width=40
    libro.sheets(mes_deb).range('E:E').column_width=12
    libro.sheets(mes_deb).range('F:F').column_width=12
    libro.sheets(mes_deb).range('G:G').column_width=12
    libro.sheets(mes_deb).range('H:H').column_width=12
    libro.sheets(mes_deb).range('I:I').column_width=14
    libro.sheets(mes_deb).range('J:J').column_width=30
    libro.sheets(mes_deb).range('E:E').number_format = '0,00'
    libro.sheets(mes_deb).range('F:F').number_format = '0,00'
    libro.sheets(mes_deb).range('G:G').number_format = '0,00'
    libro.sheets(mes_deb).range('H:H').number_format = '0,00'
    libro.sheets(mes_deb).range('I:I').number_format = '0,00'
    
    # =============================================================================================
    # GENERAMOS PLANILLA CREDITOS
    # ====> Genero el DataFrame
    c = Movi.loc[Movi["Monto"] > 0, :]
    creditos = c.copy()

    creditos['Haberes'] = np.nan
    creditos['Haberes'] = creditos.loc[creditos['Concepto'].str.contains('Acreditaci.n.+Haberes.+', regex=True, na=False), 'Monto']

    creditos['Extras']  = np.nan
    creditos['Extras']  = creditos.loc[~creditos['Concepto'].str.contains('Acreditaci.n.+Haberes.+', regex=True, na=False), 'Monto'] # ~ Todo lo que NO es 'Acreditaci.n.+Haberes.+' 
    
    #Credito por pago de plazo fijo | Crédito operación cambio ME por Plataforma Digital
    creditos['Inversiones'] = np.nan
    creditos['Inversiones'] = creditos.loc[creditos['Concepto'].str.contains('Cr.dito.+plazo\s+fijo', regex=True, na=False) | \
                                           creditos['Concepto'].str.contains('Cr.dito.+cambio\sME.+', regex=True, na=False), 'Extras']

    creditos.loc[creditos['Concepto'].str.contains('Cr.dito.+plazo\s+fijo', regex=True, na=False), 'Extras'] = np.nan 
    creditos.loc[creditos['Concepto'].str.contains('Cr.dito.+cambio\sME.+', regex=True, na=False), 'Extras'] = np.nan

    del creditos['Monto']
    del creditos['NroComprobante']

    creditos_ord = creditos.sort_values(by=['Haberes', 'Extras', 'Inversiones'], ascending=True)

    #v_suma_haberes = creditos['Haberes'].sum()
    #v_suma_extras  = creditos['Extras'].sum()
    # creditos.loc[creditos.index.max()+1, :] = [np.nan, np.nan, np.nan, v_suma_haberes, v_suma_extras]
    
    #v_suma_ingresos = v_suma_haberes + v_suma_extras

    # =====> Genero Planilla
    mes_cre = Mes + '_Creditos'
    libro.sheets.add(name=mes_cre, after=libro.sheets(mes_deb))
    libro.sheets(mes_cre).range('A1').options(pd.DataFrame, numbers='0,00', dates=dt.date, expand='table').value = creditos_ord
    table = libro.sheets(mes_cre).tables.add(source=libro.sheets(mes_cre).range('A1').expand(), name=mes_cre)
    libro.sheets(mes_cre).autofit()
    libro.sheets(mes_cre).range('A:A').column_width=3 # indice
    libro.sheets(mes_cre).range('B:B').column_width=10 # Fecha
    libro.sheets(mes_cre).range('C:C').column_width=30 # Concepto 
    libro.sheets(mes_cre).range('D:D').column_width=40  # Descripcion
    libro.sheets(mes_cre).range('E:E').column_width=12  # Haberes
    libro.sheets(mes_cre).range('F:F').column_width=12  # Extras
    libro.sheets(mes_cre).range('G:G').column_width=12  # Inversiones
    libro.sheets(mes_cre).range('E:E').number_format = '0,00'
    libro.sheets(mes_cre).range('F:F').number_format = '0,00'
    libro.sheets(mes_cre).range('G:G').number_format = '0,00'

    # =============================================================================================
    # GENERAMOS PLANILLA EFECTIVO
    # ====> Genero el DataFrame
    # 
    v_fecha = Movi.loc[0, 'Fecha'].strftime('%Y-%m-01')
    efectivo = [['Fecha', 'Concepto', 'Descripcion', 'Gastos', 'Tarjetas', 'Impuestos', 'Inversiones', 'Extras', 'Comentarios'],
                [v_fecha , '', '', 0.00, 0.00, 0.00, 0.00, 0.00, '']]

    # =====> Genero Planilla 
    mes_efe = Mes + '_Efectivo'
    libro.sheets.add(name=mes_efe, after=libro.sheets(mes_cre))
    libro.sheets(mes_efe).range('A1').options(numbers='0,00', dates=dt.date, expand='table').value = efectivo
    table = libro.sheets(mes_efe).tables.add(source=libro.sheets(mes_efe).range('A1').expand(), name=mes_efe)
    libro.sheets(mes_efe).autofit()
    #libro.sheets(mes_efe).range('A:A').column_width=3 # indice
    libro.sheets(mes_efe).range('A:A').column_width=10  # Fecha
    libro.sheets(mes_efe).range('B:B').column_width=30  # Concepto 
    libro.sheets(mes_efe).range('C:C').column_width=40  # Descripcion
    libro.sheets(mes_efe).range('D:D').column_width=12  # Gastos
    libro.sheets(mes_efe).range('E:E').column_width=12  # Tarjetas
    libro.sheets(mes_efe).range('F:F').column_width=12  # Impuestos
    libro.sheets(mes_efe).range('G:G').column_width=12  # Inversiones
    libro.sheets(mes_efe).range('H:H').column_width=12  # Extras
    libro.sheets(mes_efe).range('I:I').column_width=40  # Comentarios
    libro.sheets(mes_efe).range('D:D').number_format = '0,00'
    libro.sheets(mes_efe).range('E:E').number_format = '0,00'
    libro.sheets(mes_efe).range('F:F').number_format = '0,00'
    libro.sheets(mes_efe).range('G:G').number_format = '0,00'
    libro.sheets(mes_efe).range('H:H').number_format = '0,00'

    #libro.save(path=r'C:\Users\debus\GestionGastos\GestionGastos.xlsm') # guardar

def balance():
    #
    # CALCULA PLANILLA "Gastos" y "Saldo Mensual"
    # 
    import calendar
    libro = xw.Book.caller()
    periodo = libro.sheets('Menu').range('H2').value
    mes = periodo[0:2]
    anio = periodo[2:]

    Mes = {'01':'Ene', '02':'Feb', '03':'Mar', '04':'Abr', '05':'May', '06':'Jun', '07':'Jul', '08':'Ago', '09':'Sep', '10':'Oct', '11':'Nov', '12':'Dic' }
    
    # =============================================================================================
    # LEER PLANILLA MES_DEBITOS 
    # 
    # NOTA: La primera columna del rango se toma como indice del DataFrame, por tanto, si no hay valor de celda(indice) no trae la fila
    #
    # NO ACTUALIZAR TOTALES PLANILLA DEBITO
     
    mes_deb = Mes[mes] + '_Debitos'

    # =====> LEO PLANILLA
    debitos = libro.sheets(mes_deb).range('B1').options(pd.DataFrame, expand='table').value
    v_suma_gastos_deb       = debitos['Gastos'].sum()
    v_suma_tarjetas_deb     = debitos['Tarjetas'].sum()
    v_suma_impuestos_deb    = debitos['Impuestos'].sum()
    v_suma_inversiones_deb  = debitos['Inversiones'].sum()
    v_suma_extracciones_deb = debitos['Extracciones'].sum()

    #imp_deb_aut_nar   = debitos.loc[debitos['Descripcion'].str.contains('TARJNARANJA', regex=False, na=False), 'Impuestos'] # Serie  
    #v_imp_deb_aut_nar = imp_deb_aut_nar.iloc[0]  # Impuestos pagados con Debito Automatico Naranja
    v_suma_tarjetas_sin_imp = v_suma_tarjetas_deb # Tarjetas Sin Impuestos

    #anio = debitos.index.max().strftime('%Y')
    #mes = debitos.index.max().strftime('%m')
    # Obtengo el Ultimo dia del Mes
    ult_dia = str(calendar.monthrange(int(anio), int(mes))[1])
    v_Fecha = anio + '-' + mes + '-' + ult_dia 
    totales_deb = pd.DataFrame(columns = ['Concepto', 'Descripcion', 'Gastos', 'Tarjetas', 'Impuestos', 'Inversiones', 'Extracciones', 'Comentarios'])
    totales_deb.loc[v_Fecha, :] = ['Total :', '', v_suma_gastos_deb, v_suma_tarjetas_sin_imp, v_suma_impuestos_deb, v_suma_inversiones_deb, v_suma_extracciones_deb, '']
    

    # Obtener el rango y sumarle un offset de 4 filas
    #fila = rango.get_address(row_absolute=False, column_absolute=False)[4:]
    rango = libro.sheets(mes_deb).tables(mes_deb).data_body_range
    filas  = rango.rows.count
    fila_ini = str(filas + 4)
    rng_ini  = 'B'+ fila_ini
    libro.sheets(mes_deb).range(rng_ini).options(pd.DataFrame, expand='table').value = totales_deb


    # =============================================================================================
    # LEER PLANILLA MES_CREDITOS
    #
    # NO ACTUALIZAR TOTALES PLANILLA CREDITO
    #
    mes_cre = Mes[mes] + '_Creditos'
    creditos = libro.sheets(mes_cre).range('B1').options(pd.DataFrame, expand='table').value

    v_suma_haberes_cre = creditos['Haberes'].sum()
    v_suma_extras_cre  = creditos['Extras'].sum()
    v_suma_inversiones_cre  = creditos['Inversiones'].sum()

    totales_cre = pd.DataFrame(columns = ['Concepto', 'Descripcion', 'Haberes', 'Extras', 'Inversiones'])
    totales_cre.loc[v_Fecha, :] = ['Total :', '', v_suma_haberes_cre, v_suma_extras_cre, v_suma_inversiones_cre]


    # Obtener el rango y sumarle un offset de 3 filas
    #fila = rango.get_address(row_absolute=False, column_absolute=False)[4:]
    rango = libro.sheets(mes_cre).tables(mes_cre).data_body_range
    filas = rango.rows.count
    fila_ini = str(filas + 4)
    rng_ini  = 'B'+ fila_ini
    libro.sheets(mes_cre).range(rng_ini).options(pd.DataFrame, expand='table').value = totales_cre


    # =============================================================================================
    # LEER PLANILLA MES_EFECTIVO
    #  NO ACTUALIZAR PLANILLA EFECTIVO
    #
    mes_efe = Mes[mes] + '_Efectivo'
    efectivo = libro.sheets(mes_efe).range('A1').options(pd.DataFrame, expand='table').value

    v_suma_gastos_efe       = efectivo['Gastos'].sum()
    v_suma_tarjetas_efe     = efectivo['Tarjetas'].sum()
    v_suma_impuestos_efe    = efectivo['Impuestos'].sum()
    v_suma_inversiones_efe  = efectivo['Inversiones'].sum()
    v_suma_extras_efe       = efectivo['Extras'].sum()

    totales_efe = pd.DataFrame(columns = ['Concepto', 'Descripcion', 'Gastos', 'Tarjetas', 'Impuestos', 'Inversiones', 'Extras', 'Comentarios'])
    totales_efe.loc[v_Fecha, :] = ['Total :', '', v_suma_gastos_efe, v_suma_tarjetas_efe, v_suma_impuestos_efe, v_suma_inversiones_efe, v_suma_extras_efe, '']

    # Obtener el rango y sumarle un offset de 3 filas
    rango = libro.sheets(mes_efe).tables(mes_efe).data_body_range
    filas = rango.rows.count
    fila_ini = str(filas + 4)
    rng_ini  = 'A'+ fila_ini
    libro.sheets(mes_efe).range(rng_ini).options(pd.DataFrame, expand='table').value = totales_efe

    
    # =============================================================================================
    #  ANULADA ESTA PARTE : PLANILLA IMPUESTOS SOLO DE REGISTRO O SEGUIMIENTO DE IMPUESTOS ANUALES
    # LEER PLANILLA IMPUESTOS 
    #
    # impuestos = libro.sheets('Impuestos').range('A3').options(pd.DataFrame, expand='table').value
    # v_suma_FEB = impuestos['FEB'].sum()
    # v_suma_MAR = impuestos['MAR'].sum()
    # v_suma_ABR = impuestos['ABR'].sum()
    # v_suma_MAY = impuestos['MAY'].sum()
    # v_suma_JUN = impuestos['JUN'].sum()
    # v_suma_JUL = impuestos['JUL'].sum()
    # v_suma_AGO = impuestos['AGO'].sum()
    # v_suma_SEP = impuestos['SEP'].sum()
    # v_suma_OCT = impuestos['OCT'].sum()
    # v_suma_NOV = impuestos['NOV'].sum()
    # v_suma_DIC = impuestos['DIC'].sum()
    # v_suma_ENE = impuestos['ENE'].sum()
    # impuestos.loc['TOTAL', :] = ['', '', '', '', v_suma_FEB, v_suma_MAR, v_suma_ABR, v_suma_MAY, v_suma_JUN, v_suma_JUL, v_suma_AGO, v_suma_SEP, v_suma_OCT, v_suma_NOV, v_suma_DIC, v_suma_ENE]

    # impuestos.loc['RENTAS', :]
    # impuestos.loc['MUNICBA', :]

      
    # =============================================================================================
    # ESCRIBIR PLANILLA GASTOS
    # GRABAR NUEVOS VALORES CORRESPONDIENTES AL MES
    # ACTUALIZAR TOTALES
    gastos = libro.sheets('Gastos').range('A2').options(pd.DataFrame, expand='table').value

    col = ult_dia + '-' + Mes[mes]

    gastos.loc['CONSUMO TARJETAS', col] = v_suma_tarjetas_sin_imp + v_suma_tarjetas_efe
    gastos.loc['CONSUMO DEBITADO', col] = v_suma_gastos_deb
    gastos.loc['EFECTIVO GASTADO', col] = v_suma_gastos_efe
    gastos.loc['CONSUMO EXTRAORDINARIO', col] = 0.00
    gastos.loc[:, 'ANUAL'] = 0.00
    gastos.loc[:, 'ANUAL'] = gastos.sum(axis=1, skipna=True)

    
    v_sum_ENE_gas = gastos['31-Ene'].sum()
    v_sum_FEB_gas = gastos['28-Feb'].sum()
    v_sum_MAR_gas = gastos['31-Mar'].sum()
    v_sum_ABR_gas = gastos['30-Abr'].sum()
    v_sum_MAY_gas = gastos['31-May'].sum()
    v_sum_JUN_gas = gastos['30-Jun'].sum()
    v_sum_JUL_gas = gastos['31-Jul'].sum()
    v_sum_AGO_gas = gastos['31-Ago'].sum()
    v_sum_SEP_gas = gastos['30-Sep'].sum()
    v_sum_OCT_gas = gastos['31-Oct'].sum()
    v_sum_NOV_gas = gastos['30-Nov'].sum()
    v_sum_DIC_gas = gastos['31-Dic'].sum()
    
    '''
    gastos['ANUAL'] = gastos['31-Ene'] + gastos['28-Feb'] + gastos['31-Mar'] + gastos['30-Abr'] + gastos['31-May'] + gastos['30-Jun'] + gastos['31-Jul'] + gastos['31-Ago'] + gastos['30-Sep'] + gastos['31-Oct'] + gastos['30-Nov'] + gastos['31-Dic']
    
    gastos.loc['TOTAL CONSUMOS', :] = [v_sum_ENE_gas, v_sum_FEB_gas, v_sum_MAR_gas, v_sum_ABR_gas, v_sum_MAY_gas, v_sum_JUN_gas, v_sum_JUL_gas, v_sum_AGO_gas, v_sum_SEP_gas, v_sum_OCT_gas, v_sum_NOV_gas, v_sum_DIC_gas]
    '''
    #libro.sheets('Gastos').tables['Gastos'].update(gastos) # LICENCIA PRO
    libro.sheets('Gastos').range('A2').options(pd.DataFrame, expand='table').value = gastos
    
    # =============================================================================================
    # ESCRIBIR PLANILLA SALDO MENSUAL
    # GRABAR NUEVOS VALORES CORRESPONDIENTES AL MES
    # ACTUALIZAR TOTALES
    #
    saldo = libro.sheets('Saldo Mensual').range('A2').options(pd.DataFrame, expand='table').value
    
    v_ingreso = v_suma_haberes_cre + v_suma_extras_cre + v_suma_extras_efe
    v_egreso = v_suma_tarjetas_sin_imp + v_suma_tarjetas_efe + v_suma_gastos_deb + v_suma_gastos_efe + v_suma_impuestos_deb + v_suma_impuestos_efe

    saldo.loc['HABERES', col] = v_suma_haberes_cre
    saldo.loc['EXTRAS', col]  = v_suma_extras_cre + v_suma_extras_efe
    saldo.loc['TOTAL INGRESOS', col] = v_ingreso

    saldo.loc['CONSUMO', col]   = v_suma_tarjetas_sin_imp + v_suma_tarjetas_efe + v_suma_gastos_deb + v_suma_gastos_efe
    saldo.loc['IMPUESTOS', col] = v_suma_impuestos_deb + v_suma_impuestos_efe
    saldo.loc['TOTAL EGRESOS', col] = v_egreso

    saldo.loc['AHORRO / DEFICIT', col] = v_ingreso + v_egreso

    saldo.loc[:, 'ANUAL'] = 0.00
    saldo.loc[:, 'ANUAL'] = saldo.sum(axis=1, skipna=True)

    libro.sheets('Saldo Mensual').range('A2').options(pd.DataFrame, expand='table').value = saldo

if __name__ == "__main__":
    # xw.Book("GestionGastos.xlsm").set_mock_caller()
    main()
    