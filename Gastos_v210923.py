#
# Importar las librerías
#

import xlwings  as xw
import numpy    as np
import pandas   as pd
from datetime import date, datetime as dt, timedelta
import os
import sys
from pathlib import Path
import re
import calendar
from decimal  import Decimal
import pdfplumber

# Gráficos
# ==============================================================================
#import matplotlib.pyplot as plt
#import seaborn as sns
#%matplotlib inline
#plt.style.use('fivethirtyeight')

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
pd.set_option('display.max_colwidth', None)
pd.set_option('display.float_format', '{:.2f}'.format)



@xw.func
def hello(name):
    return f"Hello {name}!"

def main():
    #
    #
    # GESTION GASTOS  Fecha: 18/09/2023 07:00
    #
    #
    '''
    # -*- coding: utf-8 -*-
    Created on Martes 8 Marzo 2022 

    @author: Daniel Busso 
    '''
    # GENERAMOS PLANILLA MOVIMIENTOS DEL MES (POR AHORA SE LLAMA Febrero)
    #  
    # EMPIEZO PROCESANDO UNA CARPETA
    # 
 
    #libro = xw.Book('GGastos_Ago_v150923.xlsm')
    libro = xw.Book.caller()
    periodo = libro.sheets('Menu').range('H2').value
    mes  = periodo[0:2]
    anio = periodo[2:]

    # GUARDAR EL ARCHIVO CA $ 900 0020382809-Movimientos.xls COMO Movi.xlsx
    # Archivo = carpeta_trabajo + '/' + 'Movi.xls'  HAY QUE VERLO

    # NO PUEDO USAR decimal=',', thousands='.' debido al formato que tiene la columna "Monto" Por ejemplo $1.500,45
    # datos = pd.read_excel(Archivo, names=['Fecha', 'Concepto', 'Descripcion', 'Monto', 'NroComprobante'], 
                        # skiprows=5, skipfooter=2, decimal=',', thousands='.', engine='openpyxl')
    
    #=========================================================================================================================
    #
    # LEER ARCHIVO DE MOVIMIENTOS CA PESOS BANCOR y GENERAMOS EL DATAFRAME DEL MES 
    #
  
    carpeta = Path(r'G:\Mi unidad\Mis_Pagos', anio, periodo)
    Archivo = list(carpeta.glob('CA*$*0020382809*.xls')) # OJO DEBE HABER SOLO UN ARCHIVO POR CARPETA

    datos = pd.read_excel(Archivo[0], names=['Fecha', 'Concepto', 'Descripcion', 'Monto', 'NroComprobante'], skiprows=5, skipfooter=2, engine='xlrd')  
    
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

    datos_ord = Movi.sort_values(by='Fecha', ascending=True)

    #Mes = Movi.loc[0, 'Fecha'].strftime('%B')  # Febrero

    #=========================================================================================================================================== 
    #
    # ESCRIBIMOS LOS MOVIMIENTOS BANCOR A LA PLANILLA "MESES" 
    #planilla = xw.Book.caller().sheets[0]

    # CALCULAR RANGO
    # Pasar df a excel Sin Cabecera 
    rango = libro.sheets('Meses').tables('Meses').data_body_range
    filas = rango.rows.count
    fila  = str(filas + 2)
    rango_i = 'A' + fila
    libro.sheets('Meses').range(rango_i).options(pd.DataFrame, numbers='0,00', dates=dt.date, expand='table', header=0).value = datos_ord
    #table = libro.sheets('Meses').tables.add(source=libro.sheets('Meses').range('A1').expand(), name='Meses')


    libro.sheets('_xlwings.conf').visible=0  # 0 OCULTA

    # FORMATEAR PLANILLA 
    libro.sheets('Meses').autofit()
    libro.sheets('Meses').range('A:A').column_width=3   # indice
    libro.sheets('Meses').range('B:B').column_width=10  # Fecha
    libro.sheets('Meses').range('C:C').column_width=30  # Concepto
    libro.sheets('Meses').range('D:D').column_width=40  # Descripcion
    libro.sheets('Meses').range('E:E').column_width=12  # Monto
    libro.sheets('Meses').range('F:F').column_width=10  # NroComprobante
    libro.sheets('Meses').range('E:E').number_format = '0,00'    

    

    #===============================================================================================================================================
    #
    #
    #
    # GENERAMOS PLANILLA DEBITOS Debitos    ====> Genero el DataFrame
    #
    #
    d = Movi.loc[Movi["Monto"] <= 0, :]
    debitos = d.copy()

    debitos['Tarjetas'] = np.nan

    # UNA FORMA: ES CREAR PRIMERO LA COLUMNA Y DESPEJAR EL VALOR Y SU INDICE Y GUARDARLO USANDO LOC
    #tarjeta = debitos.loc[debitos['Descripcion'].str.contains('TARJNARANJA.+', regex=True, na=False), 'Monto'] # QUE PASA SI TRAE MAS DE 1 REGISTRO
    #debitos.loc[tarjeta.index, 'Tarjetas'] = tarjeta
    #debitos.loc[tarjeta.index, 'Monto']    = np.nan

    # OTRA FORMA MAS SEGURA
    # ES SEPARAR CADA CONCEPTO EN UNA COLUMNA UNICA Y LUEGO SUMAR O COMBINAR ESAS COLUMNAS EN UNA SOLA
    # NO DA ERROR SI LA BUSQUEDA TRAE MAS DE UN REGISTRO

    debitos['Tarjetas'] = debitos.loc[debitos['Descripcion'].str.contains('20175349813.+', regex=True, na=False) | \
                                        debitos['Descripcion'].str.contains('TARJETA.+', regex=True, na=False) | \
                                        debitos['Descripcion'].str.contains('VISA.+', regex=True, na=False), 'Monto']

    debitos.loc[debitos['Descripcion'].str.contains('20175349813.+', regex=True, na=False), 'Monto'] = np.nan
    debitos.loc[debitos['Descripcion'].str.contains('TARJETA.+', regex=True, na=False), 'Monto'] = np.nan
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


    # Ultimos arreglos
    del debitos['NroComprobante']
    debitos['Comentarios'] = ''
    debitos = debitos.rename(columns={'Monto':'Gastos'})
    debitos['Gastos']       = (-1) * debitos['Gastos']
    debitos['Tarjetas']     = (-1) * debitos['Tarjetas']
    debitos['Impuestos']    = (-1) * debitos['Impuestos']
    debitos['Inversiones']  = (-1) * debitos['Inversiones']
    debitos['Extracciones'] = (-1) * debitos['Extracciones']

    debitos_ord = debitos.sort_values(by=['Gastos', 'Tarjetas', 'Inversiones', 'Extracciones'], ascending=True)

    # AGREGO ULTIMA LINEA PARA DESCONTAR IMPUESTOS PAGADOS CON LA TARJETA NARANJA U OTRO RECORDATORIO
    # Obtengo el Ultimo dia del Mes
    #ult_dia = str(calendar.monthrange(int(anio), int(mes))[1])
    #v_Fecha = anio + '-' + mes + '-' + ult_dia 

    #debitos_ord.loc[900, :] = [v_Fecha, 'Impuestos debitados en Tarjeta Naranja', 'Descontar Impuestos de la Tarjeta', '', '', 0.0, '', '', '']     
    # debitos.loc[debitos.index.max()+1, :] = [np.nan, np.nan, np.nan, v_suma_gastos, v_suma_tarjetas, v_suma_impuestos, v_suma_inversiones, v_suma_extracciones, '']

    planilla = 'Debitos'
    # Calcular el Rango
    # Pasar df a excel Sin Cabecera 
    rango = libro.sheets(planilla).tables(planilla).data_body_range
    filas = rango.rows.count
    fila  = str(filas + 2)
    rango_i = 'A' + fila
    libro.sheets(planilla).range(rango_i).options(pd.DataFrame, numbers='0,00', dates=dt.date, expand='table', header=0).value = debitos_ord
    #table = libro.sheets(planilla).tables.add(source=libro.sheets(planilla).range('A1').expand(), name=planilla)

    # =====> CONFIGURO Planilla
    #mes = Movi.loc[0, 'Fecha'].strftime('%b')  # Feb 
    libro.sheets(planilla).autofit()
    libro.sheets(planilla).range('A:A').column_width=3
    libro.sheets(planilla).range('B:B').column_width=10
    libro.sheets(planilla).range('C:C').column_width=40
    libro.sheets(planilla).range('D:D').column_width=40
    libro.sheets(planilla).range('E:E').column_width=12
    libro.sheets(planilla).range('F:F').column_width=12
    libro.sheets(planilla).range('G:G').column_width=12
    libro.sheets(planilla).range('H:H').column_width=12
    libro.sheets(planilla).range('I:I').column_width=14
    libro.sheets(planilla).range('J:J').column_width=30
    libro.sheets(planilla).range('E:E').number_format = '0,00'
    libro.sheets(planilla).range('F:F').number_format = '0,00'
    libro.sheets(planilla).range('G:G').number_format = '0,00'
    libro.sheets(planilla).range('H:H').number_format = '0,00'
    libro.sheets(planilla).range('I:I').number_format = '0,00'

    # RESALTAR ULTIMA LINEA
    # 
    #rango = libro.sheets(planilla).tables(planilla).data_body_range
    #filas = rango.rows.count
    #fila  = str(filas + 1)
    #rango_i = 'A' + fila
    #rango_f = 'G' + fila
    #rango = rango_i + ':' + rango_f
    #libro.sheets(planilla).range(rango).color = (255, 255, 0)
    
    

    # ===================================================================================================================================
    #
    #
    # GENERAMOS PLANILLA CREDITOS Creditos
    # ====> Genero el DataFrame
    #
    #

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
    creditos['Comentarios'] = ''

    del creditos['Monto']
    del creditos['NroComprobante']

    creditos_ord = creditos.sort_values(by=['Haberes', 'Extras', 'Inversiones'], ascending=True)

    #v_suma_haberes = creditos['Haberes'].sum()
    #v_suma_extras  = creditos['Extras'].sum()
    # creditos.loc[creditos.index.max()+1, :] = [np.nan, np.nan, np.nan, v_suma_haberes, v_suma_extras]

    #v_suma_ingresos = v_suma_haberes + v_suma_extras

    # =====> Genero Planilla
    planilla = 'Creditos'
    # Calcular el Rango
    # Pasar df a excel Sin Cabecera 
    rango = libro.sheets(planilla).tables(planilla).data_body_range
    filas = rango.rows.count
    fila  = str(filas + 2)
    rango_i = 'A' + fila
    libro.sheets(planilla).range(rango_i).options(pd.DataFrame, numbers='0,00', dates=dt.date, expand='table', header=0).value = creditos_ord
    #table = libro.sheets(planilla).tables.add(source=libro.sheets(planilla).range('A1').expand(), name=planilla)

    # CONFIGURAMOS PLANILLA
    #     
    libro.sheets(planilla).autofit()
    libro.sheets(planilla).range('A:A').column_width=3 # indice
    libro.sheets(planilla).range('B:B').column_width=10 # Fecha
    libro.sheets(planilla).range('C:C').column_width=30 # Concepto 
    libro.sheets(planilla).range('D:D').column_width=40  # Descripcion
    libro.sheets(planilla).range('E:E').column_width=12  # Haberes
    libro.sheets(planilla).range('F:F').column_width=12  # Extras
    libro.sheets(planilla).range('G:G').column_width=12  # Inversiones
    libro.sheets(planilla).range('E:E').number_format = '0,00'
    libro.sheets(planilla).range('F:F').number_format = '0,00'
    libro.sheets(planilla).range('G:G').number_format = '0,00'

    
    
    # 
    #
    # RESALTAR LINEA CON ACREDITACION DE PLAZO FIJO
    #
    #

    fila_max = creditos_ord.index.size + 2
    creditos_ord['Fila'] = np.arange(2, fila_max, 1)

    if creditos_ord.Inversiones.any() == True:
        ind_inv_i = creditos_ord[creditos_ord["Inversiones"].notna() == True].index[0]
        ind_inv_f = creditos_ord[creditos_ord["Inversiones"].notna() == True].index.max()
        fila_i = creditos_ord.loc[ind_inv_i, 'Fila']
        fila_f = creditos_ord.loc[ind_inv_f, 'Fila']
        rango_i = 'A' + str(fila_i)
        rango_f = 'G' + str(fila_f)
        rango = rango_i + ':' + rango_f
        libro.sheets(planilla).range(rango).color = (255, 255, 0)



def balance():
    #
    #
    # CALCULA PLANILLA "Gastos" y "Saldo Mensual"
    # 
    libro = xw.Book.caller()
    #libro = xw.Book('GGastos_Ago_v150923.xlsm')
    periodo = libro.sheets('Menu').range('H2').value
    mes  = periodo[0:2]
    anio = periodo[2:]
    Mes = {'01':'Ene', '02':'Feb', '03':'Mar', '04':'Abr', '05':'May', '06':'Jun', '07':'Jul', '08':'Ago', '09':'Sep', '10':'Oct', '11':'Nov', '12':'Dic' }

    
    # ====================================================================================================================================================
    #
    # SUMAR PLANILLA Debitos
    # 
    # NOTA: La primera columna del rango se toma como indice del DataFrame, por tanto, si no hay valor de celda(indice) no trae la fila
    #
    planilla = 'Debitos'
    debitos = libro.sheets(planilla).range('A1').options(pd.DataFrame, expand='table').value
    debitos['mes']  = debitos.Fecha.dt.month
    debitos['anio'] = debitos.Fecha.dt.year

    # FILTRO POR EL MES
    mes_sum  = int(mes)
    anio_sum = int(anio)
    #efectivo_fil_mes = efectivo[efectivo['mes'] == mes_sum]
    debitos_fil_mes = debitos[(debitos['mes'] == mes_sum) & (debitos['anio'] == anio_sum)]

    v_suma_gastos_deb       = debitos_fil_mes['Gastos'].sum()
    v_suma_tarjetas_deb     = debitos_fil_mes['Tarjetas'].sum()
    #v_suma_impuestos_deb    = debitos_fil_mes['Impuestos'].sum()   # NO SE USA,  A LOS FINES DE DISCRIMAR DE GASTOS COMUNES, SON LOS IMPUESTOS DEBITADOS O PAGADOS POR RED LINK POR Ej. Rentas
    #v_suma_inversiones_deb  = debitos_fil_mes['Inversiones'].sum()  # NO SE USA,  A LOS FINES DE DISCRIMAR DE GASTOS COMUNES 
    v_suma_extracciones_deb = debitos_fil_mes['Extracciones'].sum() # NO SE USA, A LOS FINES DE DISCRIMAR DE GASTOS COMUNES 

    # v_tarjetas_sin_imp      = v_suma_tarjetas_deb - v_suma_impuestos_deb

    
    
    # ========================================================================================================================================
    #
    # SUMAR TOTALES PLANILLA Creditos
    #
    planilla = 'Creditos'
    creditos = libro.sheets(planilla).range('A1').options(pd.DataFrame, expand='table').value
    creditos['mes']  = creditos.Fecha.dt.month
    creditos['anio'] = creditos.Fecha.dt.year


    mes_sum  = int(mes)
    anio_sum = int(anio)
    creditos_fil_mes = creditos[(creditos['mes'] == mes_sum) & (creditos['anio'] == anio_sum)]

    v_suma_haberes_cre      = creditos_fil_mes['Haberes'].sum()
    v_suma_extras_cre       = creditos_fil_mes['Extras'].sum()

    # CALCULAR PLAZO FIJO
    # Cuidado de no sumar toda la acreditacion del Plazo Fijo
    # INTERES GANADO = Monto Acreditado - Deposito
    v_suma_inversiones_cre  = creditos_fil_mes['Inversiones'].sum()   


    
    # ===============================================================================================================================================
    # SUMAR PLANILLA GASTOS ON LINE (EFECTIVO) 
    #
    # REGISTRO TODO LO QUE NO ESTA BANCARIZADO: Gastos y Extras
    #
    #
    planilla = 'Registro_web'  
    reg_web = libro.sheets(planilla).range('A1').options(pd.DataFrame, expand='table').value
    reg_web['mes']  = reg_web.index.month
    reg_web['anio'] = reg_web.index.year

    # FILTRO POR ANIO - MES - MEDIO DE PAGO (EFECTIVO)
    reg_web_mes = reg_web[(reg_web['mes'] == mes_sum) & (reg_web['anio'] == anio_sum) & (reg_web['Medio de Pago'] == 'Efectivo')]

    # FILTRO POR CATEGORIA
    # GASTOS EN EFECTIVO
    reg_web_mes_gasto = reg_web_mes[(reg_web_mes['Concepto'] == 'Comida / Vestido') | (reg_web_mes['Concepto'] == 'Transporte / Mecanica') | \
                                    (reg_web_mes['Concepto'] == 'Ayudas Economicas / Donaciones') | (reg_web_mes['Concepto'] == 'Salud / Farmacia') | \
                                    (reg_web_mes['Concepto'] == 'Entretenimiento / Vacaciones / Viajes') | (reg_web_mes['Concepto'] == 'Mantenimiento / Construccion') | \
                                    (reg_web_mes['Concepto'] == 'Construccion / Mantenimiento')]

    v_suma_gastos_efe = reg_web_mes_gasto['Monto'].sum()

    #  NO SE PAGA RESUMEN DE TARJETA EN EFECTIVO
    #reg_web_mes_tarjeta  = reg_web_mes[(reg_web_mes['Concepto'] == 'Pago de Resumen de Tarjetas')]
    #v_suma_tarjetas_efe  = reg_web_mes_tarjeta['Monto'].sum()

    # NO SE PAGAN IMPUESTOS EN EFECTIVO
    # YA QUEDA REGISTRADO EN LA TABLA IMPUESTOS
    #reg_web_mes_impuesto  = reg_web_mes[(reg_web_mes['Concepto'] == 'Pago de Impuestos')]  
    #v_suma_impuestos_efe  = reg_web_mes_impuesto['Monto'].sum()

    # SE REGISTRA LA COMPRA DE DOLARES EN CAJA DE SEGURIDAD
    #reg_web_mes_inversion  = reg_web_mes[(reg_web_mes['Concepto'] == 'Compra de Dolares')]
    #v_suma_inversiones_efe = reg_web_mes_inversion['Monto'].sum()

    # EXTRAS / TRABAJOS COBRADOS
    reg_web_mes_extra  = reg_web_mes[(reg_web_mes['Concepto'] == 'Extras / Trabajos cobrados')]
    v_suma_extras_efe  = reg_web_mes_extra['Monto'].sum()

    
    
    # ============================================================================================================================
    # SUMAR PLANILLA IMPUESTOS  -----  REGISTRO O SEGUIMIENTO DE IMPUESTOS ANUALES
    # LEER PLANILLA IMPUESTOS 
    #

    planilla = 'Impuestos'  
    impuestos = libro.sheets(planilla).range('A1').options(pd.DataFrame, expand='table').value
    impuestos['mes']  = impuestos.index.month
    impuestos['anio'] = impuestos.index.year

    # FILTRO POR ANIO - MES (TODOS LOS IMPUESTOS PAGADOS MENSUALES)
    impuestos_mes = impuestos[(impuestos['mes'] == mes_sum) & (impuestos['anio'] == anio_sum)]

    # MONTO TOTAL DE IMPUESTOS MENSUALES
    v_suma_impuestos = impuestos_mes['MONTO'].sum()



    # FILTRO POR ANIO - MES - PAGADOS CON TARJETA NARANJA
    impuestos_mes_naranja = impuestos[(impuestos['mes'] == mes_sum) & (impuestos['anio'] == anio_sum) & (impuestos['MEDIO DE PAGO'] == 'NARANJA')]

    # MONTO DE IMPUESTOS PAGADOS CON TARJETA NARANJA
    v_suma_imp_naranja = impuestos_mes_naranja['MONTO'].sum()

    # MONTO TARJETAS SIN IMPUESTOS
    v_tarjetas_sin_imp = v_suma_tarjetas_deb - v_suma_imp_naranja  ### NO SE USA


    
    # ================================================================================================================================
    #
    #  TODO EL GASTO = CONSUMO + IMPUESTOS
    #
    # ESCRIBIR TABLA GASTOS_PBI
    # GRABAR NUEVOS VALORES CORRESPONDIENTES AL MES
    # 

    #gastos_pbi = pd.DataFrame(columns = ['Fecha', 'Concepto', 'Monto', 'mes', 'anio'])

    ult_dia = str(calendar.monthrange(int(anio), int(mes))[1])
    v_Fecha = anio + '-' + mes + '-' + ult_dia 
    #totales_deb.loc[999.0, :] = [v_Fecha, 'Total: ', 'Suma de Columnas: ', v_suma_gastos_deb, v_suma_tarjetas_deb, v_suma_impuestos_deb, v_suma_inversiones_deb, v_suma_extracciones_deb, '']

    # Configurar el mismo nombre de la Tabla en Excel
    planilla = 'Gastos_pbi'
    rango  = libro.sheets(planilla).tables(planilla).data_body_range
    filas  = rango.rows.count
    filas  = filas + 2
    rng_ini  = 'A'+ str(filas)

    libro.sheets(planilla).range(rng_ini).value = [v_Fecha, v_Fecha, 'CONSUMO TARJETAS', v_suma_tarjetas_deb , mes_sum, anio_sum]

    filas    = filas + 1
    rng_ini  = 'A'+ str(filas)
    libro.sheets(planilla).range(rng_ini).value = [v_Fecha, v_Fecha, 'CONSUMO DEBITADO', v_suma_gastos_deb,   mes_sum, anio_sum]

    filas    = filas + 1
    rng_ini  = 'A' + str(filas)
    libro.sheets(planilla).range(rng_ini).value = [v_Fecha, v_Fecha, 'EFECTIVO GASTADO', v_suma_gastos_efe,   mes_sum, anio_sum]

    
    
    
    # ==================================================================================================================================
    #
    # ESCRIBIR TABLA SALDOS_PBI
    # GRABAR NUEVOS VALORES CORRESPONDIENTES AL MES
    # 
    #

    ult_dia = str(calendar.monthrange(int(anio), int(mes))[1])
    v_Fecha = anio + '-' + mes + '-' + ult_dia 
    #totales_deb.loc[999.0, :] = [v_Fecha, 'Total: ', 'Suma de Columnas: ', v_suma_gastos_deb, v_suma_tarjetas_deb, v_suma_impuestos_deb, v_suma_inversiones_deb, v_suma_extracciones_deb, '']

    # Configurar el mismo nombre de la Tabla en Excel
    planilla = 'Saldos_pbi'   
    rango  = libro.sheets(planilla).tables(planilla).data_body_range
    filas  = rango.rows.count


    filas  = filas + 2
    rng_ini  = 'A'+ str(filas)
    libro.sheets(planilla).range(rng_ini).value = [v_Fecha, v_Fecha, 'HABERES', v_suma_haberes_cre , mes_sum, anio_sum]

    v_extras = v_suma_extras_cre + v_suma_inversiones_cre + v_suma_extras_efe
    filas    = filas + 1
    rng_ini  = 'A'+ str(filas)
    libro.sheets(planilla).range(rng_ini).value = [v_Fecha, v_Fecha, 'EXTRAS', v_extras,   mes_sum, anio_sum]

    v_gasto   = v_suma_gastos_deb + v_suma_tarjetas_deb + v_suma_gastos_efe
    v_consumo = v_gasto - v_suma_impuestos
    filas    = filas + 1
    rng_ini  = 'A' + str(filas)
    libro.sheets(planilla).range(rng_ini).value = [v_Fecha, v_Fecha, 'CONSUMO', v_consumo,   mes_sum, anio_sum]

    filas    = filas + 1
    rng_ini  = 'A' + str(filas)
    libro.sheets(planilla).range(rng_ini).value = [v_Fecha, v_Fecha, 'IMPUESTOS', v_suma_impuestos,   mes_sum, anio_sum]

    
def Naranja():  
    #
    #
    #
    # GESTION GASTOS  Fecha: 14/09/2023 07:00
    '''
    # -*- coding: utf-8 -*-
    Created on Martes 8 Marzo 2022 

    @author: Daniel Busso 
    '''

    libro = xw.Book.caller()
    #libro = xw.Book('GGastos_Ago_v150923.xlsm')

    periodo = libro.sheets('Menu').range('H2').value
    mes  = periodo[0:2]
    anio = periodo[2:]
    Meses = {'01':'Ene', '02':'Feb', '03':'Mar', '04':'Abr', '05':'May', '06':'Jun', '07':'Jul', '08':'Ago', '09':'Sep', '10':'Oct', '11':'Nov', '12':'Dic' }
    Mes = Meses[mes]
    periodo_excel = Mes + '-' + anio

    carpeta = Path(r'G:\Mi unidad\Mis_Pagos', anio, periodo)  
    Archivo = list(carpeta.glob('ResumenNaranja*.pdf')) # OJO DEBE HABER SOLO UN ARCHIVO POR CARPETA
    Docupdf = Archivo[0]


    #
    #
    #
    #    LECTURA DEL RESUMEN TARJETA NARANJA
    #
    #
    #
    #

    v_msg_error = 'FORMATO ARCHIVO IMPORTACION FALLIDA'

    #pdf = pdfplumber.open(Docupdf, password='17534981')
    pdf   = pdfplumber.open(Docupdf)
    pages = pdf.pages

    text = pages[0].extract_text()
    for linea in text.split('\n'):
        if re.search(r'^y vence el .*', linea):
            fecha = linea[11:19].strip()     #  10/08/23'
            fecha = fecha.replace('/', '')  # '100823'
            mes   = fecha[2:4]              # '08'
            anio = '20' + fecha[4:]        # '2023'
            mesanio_resumen = mes + anio    # '102023'

    if (periodo == mesanio_resumen):
        for page in pages:
            text = page.extract_text()
            for linea in text.split('\n'):
                if re.search(r'^\d{2}/\d{2}/\d{2} Naranja X .*', linea):
                    f1 = linea[0:8].strip()
                    # separa de la cadena DD/MM/AA en [ 'DD', 'MM', 'AA']
                    f1 = f1.split('/')
                    f1.reverse()  # Reversa la lista ['AA', 'MM', 'DD']
                    sep = '-'
                    # Convierte la lista en cadena 'AA-MM-DD'
                    v_fecha = sep.join(f1)
                    v_fecha = '20' + v_fecha   # Convierte año a 4 digitos 'AAAA-MM-DD'
                    # v_anio    = f1[0]
                    # v_mes     = f1[1]
                    # v_mesanio = v_mes + v_anio
                    v_tarjeta = linea[9:18].strip()
                    v_cupon   = linea[19:28].strip()
                    v_detalle = linea[29:80].strip()
                    v_plan    = linea[81:97].strip()

                    # Manipulamos la columna Pesos
                    v_pesos   = linea[100:110].strip()
                    if (v_pesos == ''): 
                        v_pesos_dec = ' '
                    else:
                        v_pesos   = re.sub('[ªº\!#$%&/()_=?¿¡@]', '', v_pesos)
                        # quita todos los puntos indicadores de miles o diezmiles, etc.
                        v_pesos = v_pesos.replace('.', '')
                        # reemplaza la coma decimal por el punto decimal
                        v_pesos = v_pesos.replace(',', '.')
                        v_pesos_dec = Decimal(v_pesos)

                    # Manipulamos la Columna Dolares
                    v_dolares = linea[112:124].strip()
                    if (v_dolares == ''): 
                        v_dolares_dec = ' '
                    else:
                        v_dolares = re.sub('[ªº\!#$%&/()_=?¿¡@]', '', v_dolares)
                        v_dolares = v_dolares.replace('.', '')
                        v_dolares = v_dolares.replace(',', '.')
                        v_dolares_dec = Decimal(v_dolares)

                    # Configurar el mismo nombre de la Tabla en Excel
                    planilla = 'Naranja'
                    rango  = libro.sheets(planilla).tables(planilla).data_body_range
                    filas  = rango.rows.count
                    filas  = filas + 2
                    rng_ini  = 'A'+ str(filas)

                    libro.sheets(planilla).range(rng_ini).value = [v_fecha, periodo_excel, v_tarjeta, v_cupon, v_detalle, v_plan, v_pesos_dec, v_dolares_dec, ' ' ]

                elif re.search(r'^\d{2}/\d{2}/\d{2} NX Visa .*', linea):
                    f1 = linea[0:8].strip()
                    # separa de la cadena DD/MM/AA en [ 'DD', 'MM', 'AA']
                    f1 = f1.split('/')
                    f1.reverse()  # Reversa la lista ['AA', 'MM', 'DD']
                    sep = '-'
                    # Convierte la lista en cadena 'AA-MM-DD'
                    v_fecha = sep.join(f1)
                    v_fecha = '20' + v_fecha   # Convierte año a 4 digitos 'AAAA-MM-DD'
                    # v_anio    = f1[0]
                    # v_mes     = f1[1]
                    # v_mesanio = v_mes + v_anio
                    v_tarjeta = linea[9:18].strip()
                    v_cupon   = linea[19:28].strip()
                    v_detalle = linea[29:80].strip()
                    v_plan    = linea[81:97].strip()

                    # Manipulamos la columna Pesos
                    v_pesos   = linea[100:110].strip()
                    if (v_pesos == ''): 
                        v_pesos_dec = ' '
                    else:
                        v_pesos   = re.sub('[ªº\!#$%&/()_=?¿¡@]', '', v_pesos)
                        # quita todos los puntos indicadores de miles o diezmiles, etc.
                        v_pesos = v_pesos.replace('.', '')
                        # reemplaza la coma decimal por el punto decimal
                        v_pesos = v_pesos.replace(',', '.')
                        v_pesos_dec = Decimal(v_pesos)

                    # Manipulamos la Columna Dolares
                    v_dolares = linea[112:124].strip()
                    if (v_dolares == ''): 
                        v_dolares_dec = ' '
                    else:
                        v_dolares = re.sub('[ªº\!#$%&/()_=?¿¡@]', '', v_dolares)
                        v_dolares = v_dolares.replace('.', '')
                        v_dolares = v_dolares.replace(',', '.')
                        v_dolares_dec = Decimal(v_dolares)

                    # Configurar el mismo nombre de la Tabla en Excel
                    planilla = 'Naranja'
                    rango  = libro.sheets(planilla).tables(planilla).data_body_range
                    filas  = rango.rows.count
                    filas  = filas + 2
                    rng_ini  = 'A'+ str(filas)

                    libro.sheets(planilla).range(rng_ini).value = [v_fecha, periodo_excel, v_tarjeta, v_cupon, v_detalle, v_plan, v_pesos_dec, v_dolares_dec, ' ' ]
                else:
                    pass

    else:
        v_msg_error = 'No coincide el Mes Anio Solicitado con la Fecha del Resumen Naranja'
        v_comentario = 'IMPORTACION FALLIDA'

        planilla = 'Naranja'
        rango  = libro.sheets(planilla).tables(planilla).data_body_range
        filas  = rango.rows.count
        filas  = filas + 2
        rng_ini  = 'A'+ str(filas)

        libro.sheets(planilla).range(rng_ini).value = [v_fecha, periodo_excel,'', '', v_msg_error, v_comentario  ]

    pdf.close()

    

if __name__ == "__main__":
    # xw.Book("GestionGastos.xlsm").set_mock_caller()
    main()
    