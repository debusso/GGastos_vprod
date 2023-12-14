# Cómo importar las librerías
import xlwings  as xw
import numpy    as np
import pandas   as pd
import datetime as dt
import os
import sys
from pathlib import Path
import re
import calendar

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
    
    # ==========================================================================================================================================================================================================
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

    # ======================================================================================================================================================================================================== 
    # ESCRIBIMOS LOS MOVIMIENTOS BANCOR A LA PLANILLA "MESES" 
    #planilla = xw.Book.caller().sheets[0]

    Meses = {'01':'Ene', '02':'Feb', '03':'Mar', '04':'Abr', '05':'May', '06':'Jun', '07':'Jul', '08':'Ago', '09':'Sep', '10':'Oct', '11':'Nov', '12':'Dic' }
    Mes = Meses[mes]

    # CALCULAR RANGO
    if Mes == 'Ene' :
        rango_i = 'A1'
        # pasar df a excel
        # Crear table
        libro.sheets('Meses').range(rango_i).options(pd.DataFrame, numbers='0,00', dates=dt.date, expand='table').value = datos_ord
        #table = libro.sheets('Meses').tables.add(source=libro.sheets('Meses').range('A1').expand(), name='Meses')
    else:
        # Calcular el Rango
        # Pasar df a excel Sin Cabecera 
        rango = libro.sheets('Meses').tables('Meses').data_body_range
        filas = rango.rows.count
        fila  = str(filas + 2)
        rango_i = 'A' + fila
        libro.sheets('Meses').range(rango_i).options(pd.DataFrame, numbers='0,00', dates=dt.date, expand='table', header=0).value = datos_ord
        #table = libro.sheets('Meses').tables.add(source=libro.sheets('Meses').range('A1').expand(), name='Meses')


    libro.sheets('_xlwings.conf').visible=0  # 0 OCULTA
    #libro = xw.Book()
    #libro.sheets.add(name=Mes, after=libro.sheets(planilla)) # OJO TODAS LAS PLANILLAS SE VAN AGREGAR DESPUES DE ENERO
    
    libro.sheets('Meses').autofit()
    libro.sheets('Meses').range('A:A').column_width=3   # indice
    libro.sheets('Meses').range('B:B').column_width=10  # Fecha
    libro.sheets('Meses').range('C:C').column_width=30  # Concepto
    libro.sheets('Meses').range('D:D').column_width=40  # Descripcion
    libro.sheets('Meses').range('E:E').column_width=12  # Monto
    libro.sheets('Meses').range('F:F').column_width=10  # NroComprobante
    libro.sheets('Meses').range('E:E').number_format = '0,00'
    

    # ========================================================================================================================================================================================================
    # GENERAMOS PLANILLA DEBITOS Debitos
    # ====> Genero el DataFrame
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
    ult_dia = str(calendar.monthrange(int(anio), int(mes))[1])
    v_Fecha = anio + '-' + mes + '-' + ult_dia 

    debitos_ord.loc[900, :] = [v_Fecha, 'Impuestos debitados en Tarjeta Naranja', 'Descontar Impuestos de la Tarjeta', '', -0.0, '', '', '', '']     
    # debitos.loc[debitos.index.max()+1, :] = [np.nan, np.nan, np.nan, v_suma_gastos, v_suma_tarjetas, v_suma_impuestos, v_suma_inversiones, v_suma_extracciones, '']

    planilla = 'Debitos'
    if Mes == 'Ene' :
        rango_i = 'A1'
        # pasar df a excel
        # Crear table
        libro.sheets(planilla).range(rango_i).options(pd.DataFrame, numbers='0,00', dates=dt.date, expand='table').value = debitos_ord
        #table = libro.sheets(planilla).tables.add(source=libro.sheets(planilla).range('A1').expand(), name=planilla)
    else:
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
    rango = libro.sheets(planilla).tables(planilla).data_body_range
    filas = rango.rows.count
    fila  = str(filas + 1)
    rango_i = 'A' + fila
    rango_f = 'F' + fila
    rango = rango_i + ':' + rango_f
    libro.sheets(planilla).range(rango).color = (255, 255, 0)
        
    # ===========================================================================================================================================================================================================
    # GENERAMOS PLANILLA CREDITOS Creditos
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
    if Mes == 'Ene' :
        rango_i = 'A1'
        # pasar df a excel
        # Crear table
        libro.sheets(planilla).range(rango_i).options(pd.DataFrame, numbers='0,00', dates=dt.date, expand='table').value = creditos_ord
        #table = libro.sheets(planilla).tables.add(source=libro.sheets(planilla).range('A1').expand(), name=planilla)
    else:
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

    # RESALTAR ULTIMAS LINEAS
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

    
    # =============================================================================================
    # GENERAMOS PLANILLA EFECTIVO

    # ====> Genero el DataFrame
    # 
    # v_fecha = Movi.loc[0, 'Fecha'].strftime('%Y-%m-01')
    # efectivo = [['Fecha', 'Concepto', 'Descripcion', 'Gastos', 'Tarjetas', 'Impuestos', 'Inversiones', 'Extras', 'Comentarios'],
    #            [v_fecha , '', '', 0.00, 0.00, 0.00, 0.00, 0.00, '']]

    # =====> Genero Planilla 

    # mes_efe = Mes + '_Efectivo'
    # CAMBIAR EL NOMBRE DE LA PLANILLA DE EFECTIVO A mes_efe

    # libro.sheets('Efectivo').name = Mes + '_Efectivo'

    #libro.sheets.add(name=mes_efe, after=libro.sheets(mes_cre))
    #libro.sheets(mes_efe).range('A1').options(numbers='0,00', dates=dt.date, expand='table').value = efectivo
    # table = libro.sheets(mes_efe).tables.add(source=libro.sheets(mes_efe).range('A1').expand(), name=mes_efe)
    #libro.sheets(mes_efe).autofit()
    #libro.sheets(mes_efe).range('A:A').column_width=3 # indice
    # libro.sheets(mes_efe).range('A:A').column_width=10  # Fecha
    # libro.sheets(mes_efe).range('B:B').column_width=30  # Concepto 
    # libro.sheets(mes_efe).range('C:C').column_width=40  # Descripcion
    # libro.sheets(mes_efe).range('D:D').column_width=12  # Gastos
    # libro.sheets(mes_efe).range('E:E').column_width=12  # Tarjetas
    # libro.sheets(mes_efe).range('F:F').column_width=12  # Impuestos
    # libro.sheets(mes_efe).range('G:G').column_width=12  # Inversiones
    # libro.sheets(mes_efe).range('H:H').column_width=12  # Extras
    # libro.sheets(mes_efe).range('I:I').column_width=40  # Comentarios
    # libro.sheets(mes_efe).range('D:D').number_format = '0,00'
    # libro.sheets(mes_efe).range('E:E').number_format = '0,00'
    # libro.sheets(mes_efe).range('F:F').number_format = '0,00'
    # libro.sheets(mes_efe).range('G:G').number_format = '0,00'
    # libro.sheets(mes_efe).range('H:H').number_format = '0,00'

    #mes_efe = Mes + '_Efectivo'
    # CAMBIAR EL NOMBRE DE LA PLANILLA DE EFECTIVO A mes_efe
    #libro.sheets('Respuestas de Formulario1').name = Mes + '_Efectivo'



    #libro.save(path=r'C:\Users\debus\GestionGastos\GestionGastos.xlsm') # guardar

def balance():
    #
    # CALCULA PLANILLA "Gastos" y "Saldo Mensual"
    # 
    import calendar
    libro = xw.Book.caller()
    periodo = libro.sheets('Menu').range('H2').value
    mes  = periodo[0:2]
    anio = periodo[2:]
    Mes = {'01':'Ene', '02':'Feb', '03':'Mar', '04':'Abr', '05':'May', '06':'Jun', '07':'Jul', '08':'Ago', '09':'Sep', '10':'Oct', '11':'Nov', '12':'Dic' }
    
    # ==========================================================================================================================================================================================
    # SUMAR PLANILLA MES_DEBITOS 
    # 
    # NOTA: La primera columna del rango se toma como indice del DataFrame, por tanto, si no hay valor de celda(indice) no trae la fila
    #
    # ACTUALIZAR TOTALES PLANILLA DEBITO
     
    # =====> LEO PLANILLA
    #
    # mes_deb = Mes[mes] + '_Debitos'
    planilla = 'Debitos'
    debitos = libro.sheets(planilla).range('A1').options(pd.DataFrame, expand='table').value
    debitos['mes'] = debitos.Fecha.dt.month

    # FILTRO POR EL MES
    mes_sum = int(mes)
    #efectivo_fil_mes = efectivo[efectivo['mes'] == mes_sum]
    debitos_fil_mes = debitos[(debitos['mes'] == mes_sum)]

    v_suma_gastos_deb       = debitos_fil_mes['Gastos'].sum()
    v_suma_tarjetas_deb     = debitos_fil_mes['Tarjetas'].sum()
    v_suma_impuestos_deb    = debitos_fil_mes['Impuestos'].sum()    # A LOS FINES DE DISCRIMINAR NO SE USA PARA SUMAR
    v_suma_inversiones_deb  = debitos_fil_mes['Inversiones'].sum()  # A LOS FINES DE DISCRIMINAR NO SE USA PARA SUMAR
    v_suma_extracciones_deb = debitos_fil_mes['Extracciones'].sum() # A LOS FINES DE DISCRIMINAR NO SE USA PARA SUMAR

    #imp_deb_aut_nar   = debitos.loc[debitos['Descripcion'].str.contains('TARJNARANJA', regex=False, na=False), 'Impuestos'] # Serie  
    #v_imp_deb_aut_nar = imp_deb_aut_nar.iloc[0]  # Impuestos pagados con Debito Automatico Naranja
    
    #anio = debitos.index.max().strftime('%Y')
    #mes = debitos.index.max().strftime('%m')
    # Obtengo el Ultimo dia del Mes
    
    #debitos.loc[v_Fecha, :] = ['Impuestos debitados en la Tarjeta Naranja', '', '', 0.0, '', '', '', '']


    # =====================================================================================================================================================================================
    # SUMAR PLANILLA MES_CREDITOS
    #
    # ACTUALIZAR TOTALES PLANILLA CREDITO
    #
    planilla = 'Creditos'
    creditos = libro.sheets(planilla).range('A1').options(pd.DataFrame, expand='table').value
    creditos['mes'] = creditos.Fecha.dt.month

    mes_sum = int(mes)
    creditos_fil_mes = creditos[(creditos['mes'] == mes_sum)]

    v_suma_haberes_cre      = creditos_fil_mes['Haberes'].sum()
    v_suma_extras_cre       = creditos_fil_mes['Extras'].sum()
    v_suma_inversiones_cre  = creditos_fil_mes['Inversiones'].sum() 

    #totales_cre = pd.DataFrame(columns = ['Concepto', 'Descripcion', 'Haberes', 'Extras', 'Inversiones'])
    #creditos.loc[999.0, :] = [v_Fecha,'Total :', 'Suma de Columnas: ', v_suma_haberes_cre, v_suma_extras_cre, v_suma_inversiones_cre, '']


    # Obtener el rango y sumarle un offset de 3 filas
    #fila = rango.get_address(row_absolute=False, column_absolute=False)[4:]
    #rango = libro.sheets(mes_cre).tables(mes_cre).data_body_range
    #filas = rango.rows.count
    #fila_ini = str(filas + 4)
    #rng_ini  = 'B' + fila_ini
    #libro.sheets(planilla).range('A1').options(pd.DataFrame, expand='table').value = creditos


    # =============================================================================================================================================================================================
    # SUMAR PLANILLA EFECTIVO_GASTOS
    #
    # ACTUALIZAR PLANILLA EFECTIVO_GASTOS
    #
    
    planilla = 'Gastos_diarios'  ### GASTOS
    efectivo = libro.sheets(planilla).range('A1').options(pd.DataFrame, expand='table').value
    efectivo['mes'] = efectivo.index.month

    # FILTRO POR EL MES
    mes_sum = int(mes)
    #efectivo_fil_mes = efectivo[efectivo['mes'] == mes_sum]
    efectivo_fil_mes = efectivo[(efectivo['mes'] == mes_sum) & (efectivo['Medio de Pago'] == 'Efectivo')]

    v_suma_gastos_efe    = efectivo_fil_mes['Gastos'].sum()
    v_suma_tarjetas_efe  = efectivo_fil_mes['Tarjetas'].sum()  # DESCONTAR IMPUESTOS EN PLANILLA 
    v_suma_impuestos_efe = efectivo_fil_mes['Impuestos'].sum() # A LOS FINES DE DISCRIMINAR NO SE USA PARA SUMAR
    #v_suma_inversiones_efe  = efectivo_fil_mes['Inversiones'].sum()

    #efectivo_fil_mes.loc[v_Fecha, :] = ['Impuestos debitados en la Tarjeta', '', '', 0.0, '', '']
    #efectivo_fil_mes.loc[v_Fecha, :] = ['TOTAL:', '', v_suma_gastos_efe, v_suma_tarjetas_efe, v_suma_impuestos_efe, '']
    #libro.sheets(planilla).range('A1').options(pd.DataFrame, expand='table').value = efectivo_fil_mes



    # ===============================================================================================================================================================================================
    # SUMAR PLANILLA EFECTIVO_EXTRAS
    #  ACTUALIZAR PLANILLA EFECTIVO_EXTRAS
    #
    planilla2 = 'Efectivo_Extras'  ### GASTOS
    efectivo_ext = libro.sheets(planilla2).range('A1').options(pd.DataFrame, expand='table').value

    if efectivo_ext.index.empty:
        v_suma_extras_efe = 0.0
    else:
        efectivo_ext['mes'] = efectivo_ext.index.month

        # FILTRO POR EL MES
        mes_sum = int(mes)
        efectivo_ext_fil_mes = efectivo_ext[efectivo_ext['mes'] == mes_sum] # FILTRO EL MES
        v_suma_extras_efe = efectivo_ext_fil_mes['Extras'].sum()     # CALCULO EL MONTO TOTAL 

    #efectivo_ext_fil_mes.loc[v_Fecha, :] = [v_suma_extras_efe, 'TOTAL:', '']
    #libro.sheets(planilla2).range('A1').options(pd.DataFrame, expand='table').value = efectivo_ext_fil_mes



    #totales_efe = pd.DataFrame(columns = ['Concepto', 'Descripcion', 'Gastos', 'Impuestos', 'Tarjetas', 'Extras', 'Inversiones', 'Comentarios'])
    #totales_efe.loc[v_Fecha, :] = ['Total :', '', v_suma_gastos_efe, v_suma_impuestos_efe, v_suma_tarjetas_efe, v_suma_extras_efe, '']

    # PARA ESCRIBIR EN LA PLANILLA LOS TOTALES  --- DESPUES LO VOY A SACAR
    # Obtener el rango y sumarle un offset de 3 filas
    #rango = libro.sheets('Efectivo').tables('Efectivo').data_body_range
    #filas = rango.rows.count
    #fila_ini = str(filas + 4)
    #rng_ini  = 'A'+ fila_ini 
    #libro.sheets('Efectivo').range(rng_ini).options(pd.DataFrame, expand='table').value = totales_efe


    #====================================================================================================================================================================================
    #
    # AL FINAL ESCRIBIMOS VALORES EN LA PLANILLA TOTALES SIMILAR A LA TABLA MOVIMIENTOS EN DJANGO
    #

    totales = pd.DataFrame(columns = ['Concepto', 'Descripcion', 'Gastos', 'Tarjetas', 'Impuestos', 'Inversiones', 'Extracciones', 'Gastos_efe', 'Tarjetas_efe', 'Impuestos_efe',  'Haberes', 'Extras', 'Inversiones_acred', 'Extras_efe', 'Comentarios'])


    ult_dia = str(calendar.monthrange(int(anio), int(mes))[1])
    v_Fecha = anio + '-' + mes + '-' + ult_dia 
    #totales_deb.loc[999.0, :] = [v_Fecha, 'Total: ', 'Suma de Columnas: ', v_suma_gastos_deb, v_suma_tarjetas_deb, v_suma_impuestos_deb, v_suma_inversiones_deb, v_suma_extracciones_deb, '']
    totales.loc[v_Fecha, :] = ['Total: ', 'Suma de Columnas: ', v_suma_gastos_deb, v_suma_tarjetas_deb, v_suma_impuestos_deb, v_suma_inversiones_deb, v_suma_extracciones_deb,\
                                v_suma_gastos_efe, v_suma_tarjetas_efe, v_suma_impuestos_efe, \
                                v_suma_haberes_cre, v_suma_extras_cre, v_suma_inversiones_cre, v_suma_extras_efe,'']

    

    # Obtener el rango 
    # fila = rango.get_address(row_absolute=False, column_absolute=False)[4:]
    planilla = 'Totales'
    rango = libro.sheets(planilla).tables(planilla).data_body_range
    filas  = rango.rows.count
    fila_ini = str(filas + 2)
    rng_ini  = 'A'+ fila_ini
    libro.sheets(planilla).range(rng_ini).options(pd.DataFrame, expand='table', header=0).value = totales


    
    # ============================================================================================================================================================================================================
    # SUMAR PLANILLA IMPUESTOS  -----  SOLO DE REGISTRO O SEGUIMIENTO DE IMPUESTOS ANUALES
    # LEER PLANILLA IMPUESTOS 
    #
    impuestos_anio = 'Impuestos_' + anio
    impuestos = libro.sheets(impuestos_anio).range('A3').options(pd.DataFrame, expand='table').value

    impuestos.loc['TOTAL', :] = 0.0
    v_suma_FEB = impuestos['FEB'].sum()
    v_suma_MAR = impuestos['MAR'].sum()
    v_suma_ABR = impuestos['ABR'].sum()
    v_suma_MAY = impuestos['MAY'].sum()
    v_suma_JUN = impuestos['JUN'].sum()
    v_suma_JUL = impuestos['JUL'].sum()
    v_suma_AGO = impuestos['AGO'].sum()
    v_suma_SEP = impuestos['SEP'].sum()
    v_suma_OCT = impuestos['OCT'].sum()
    v_suma_NOV = impuestos['NOV'].sum()
    v_suma_DIC = impuestos['DIC'].sum()
    v_suma_ENE = impuestos['ENE'].sum()

    #totales_imp = pd.DataFrame(columns = ['Administracion', 'Tipo', 'Descripcion', 'Cuenta', 'Plan', 'Cuota 1', 'Cuota 2', 'Cuota 3', 'Cuota 4','Cuota 5','Cuota 6','Cuota 7','Cuota 8','Cuota 9','Cuota 10','Cuota 11','Cuota 12'])
    impuestos.loc['TOTAL', :] = ['', '', '', '', v_suma_FEB, v_suma_MAR, v_suma_ABR, v_suma_MAY, v_suma_JUN, v_suma_JUL, v_suma_AGO, v_suma_SEP, v_suma_OCT, v_suma_NOV, v_suma_DIC, v_suma_ENE]

    #impuestos.loc['RENTAS', :]
    #impuestos.loc['MUNICBA', :]

    # PARA ESCRIBIR EN LA PLANILLA LOS TOTALES  
    # Obtener el rango y sumarle un offset de 3 filas
    #rango = libro.sheets(impuestos_anio).tables(impuestos_anio).data_body_range
    #filas = rango.rows.count
    #fila_ini = str(filas + 4)
    #rng_ini  = 'A'+ fila_ini 
    libro.sheets(impuestos_anio).range('A3').options(pd.DataFrame, expand='table').value = impuestos

      
    # =========================================================================================================================================================================================================
    # ESCRIBIR PLANILLA GASTOS
    # GRABAR NUEVOS VALORES CORRESPONDIENTES AL MES
    # ACTUALIZAR TOTALES
    gastos = libro.sheets('Gastos').range('A2').options(pd.DataFrame, expand='table').value

    col = ult_dia + '-' + Mes[mes]

    MES = Mes[mes].upper()
    #v_suma_tarjetas_sin_imp = v_suma_tarjetas_deb - impuestos.loc['TOTAL', MES]   # Tarjetas Sin Impuestos    
    v_suma_tarjetas_sin_imp = v_suma_tarjetas_deb + v_suma_tarjetas_efe           # Tarjetas Sin Impuestos

    gastos.loc['CONSUMO TARJETAS', col] = v_suma_tarjetas_sin_imp
    gastos.loc['CONSUMO DEBITADO', col] = v_suma_gastos_deb
    gastos.loc['EFECTIVO GASTADO', col] = v_suma_gastos_efe
    gastos.loc['CONSUMO EXTRAORDINARIO', col] = 0.00
    gastos.loc[:, 'ANUAL'] = 0.00  # Es necesario porque sino la proxima instruccion me suma el valor de la columna ANUAL
    gastos.loc[:, 'ANUAL'] = gastos.sum(axis=1, skipna=True) # para todas las filas en la columna ANUAL coloca la suma de la fila correspondiente

    # Sumo columnas
    gastos.loc['TOTAL GASTOS', :] = 0.0
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
    v_sum_ANUAL   = gastos['ANUAL'].sum()
    gastos.loc['TOTAL GASTOS', :] = [v_sum_ENE_gas, v_sum_FEB_gas, v_sum_MAR_gas, v_sum_ABR_gas, v_sum_MAY_gas, v_sum_JUN_gas, v_sum_JUL_gas, v_sum_AGO_gas, v_sum_SEP_gas, v_sum_OCT_gas, v_sum_NOV_gas, v_sum_DIC_gas, v_sum_ANUAL]
    
    #gastos['ANUAL'] = gastos['31-Ene'] + gastos['28-Feb'] + gastos['31-Mar'] + gastos['30-Abr'] + gastos['31-May'] + gastos['30-Jun'] + gastos['31-Jul'] + gastos['31-Ago'] + gastos['30-Sep'] + gastos['31-Oct'] + gastos['30-Nov'] + gastos['31-Dic']

    
    
    #libro.sheets('Gastos').tables['Gastos'].update(gastos) # LICENCIA PRO
    libro.sheets('Gastos').range('A2').options(pd.DataFrame, expand='table').value = gastos
    
    # =================================================================================================================================================================================================================
    # ESCRIBIR PLANILLA SALDO MENSUAL
    # GRABAR NUEVOS VALORES CORRESPONDIENTES AL MES
    # ACTUALIZAR TOTALES
    #
    saldo = libro.sheets('Saldo_Mensual').range('A2').options(pd.DataFrame, expand='table').value
    
    v_ingreso = v_suma_haberes_cre + v_suma_extras_cre + v_suma_inversiones_cre + v_suma_extras_efe    # v_suma_inversiones_cre es la ganancia del plazo fijo
    v_egreso  = -v_suma_tarjetas_sin_imp - v_suma_gastos_deb - v_suma_gastos_efe - impuestos.loc['TOTAL', MES]


    # CUIDADO 
    # v_suma_inversiones_cre = Monto acreditado - Deposito

    saldo.loc['HABERES', col] = v_suma_haberes_cre
    saldo.loc['EXTRAS', col]  = v_suma_extras_cre + v_suma_extras_efe + v_suma_inversiones_cre 
    saldo.loc['TOTAL INGRESOS', col] = v_ingreso

    #saldo.loc['CONSUMO', col]   = v_suma_tarjetas_sin_imp + v_suma_tarjetas_efe + v_suma_gastos_deb + v_suma_gastos_efe
    saldo.loc['CONSUMO', col]   = -gastos.loc['TOTAL GASTOS', col] 
    saldo.loc['IMPUESTOS', col] = -impuestos.loc['TOTAL', MES]
    saldo.loc['TOTAL EGRESOS', col] = v_egreso

    saldo.loc['AHORRO / DEFICIT', col] = v_ingreso + v_egreso

    saldo.loc[:, 'ANUAL'] = 0.00
    saldo.loc[:, 'ANUAL'] = saldo.sum(axis=1, skipna=True)

    libro.sheets('Saldo_Mensual').range('A2').options(pd.DataFrame, expand='table').value = saldo

if __name__ == "__main__":
    # xw.Book("GestionGastos.xlsm").set_mock_caller()
    main()
    