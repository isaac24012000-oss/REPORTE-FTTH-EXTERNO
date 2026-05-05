import pandas as pd
from datetime import datetime

# Simular la función get_instaladas_por_semana para Abril
def get_instaladas_por_semana_debug(mes_seleccionado="Abril"):
    """Debug version"""
    df_drive = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='DRIVE')
    
    if df_drive is None or df_drive.empty:
        return pd.DataFrame()
    
    # Conversión simple, sin limpieza excesiva
    df_temp = df_drive.copy()
    df_temp['FECHA'] = pd.to_datetime(df_temp['FECHA'], errors='coerce')
    
    # Mapeo de meses
    mes_numeros = {
        'Enero': 1, 'Febrero': 2, 'Marzo': 3, 'Abril': 4,
        'Mayo': 5, 'Junio': 6, 'Julio': 7, 'Agosto': 8,
        'Septiembre': 9, 'Octubre': 10, 'Noviembre': 11, 'Diciembre': 12
    }
    
    mes_num = mes_numeros.get(mes_seleccionado, None)
    print(f"Mes seleccionado: {mes_seleccionado}, mes_num: {mes_num}")
    
    if mes_num is None:
        return pd.DataFrame()
    
    # Filtrar por fecha válida
    df_temp = df_temp[df_temp['FECHA'].notna()]
    print(f"Registros con FECHA válida: {len(df_temp)}")
    
    # FILTRO POR FECHA ACTUAL - no mostrar fechas futuras
    fecha_actual = pd.Timestamp.today()
    print(f"Fecha actual: {fecha_actual}")
    
    df_temp = df_temp[df_temp['FECHA'] <= fecha_actual]
    print(f"Registros no futuros: {len(df_temp)}")
    
    # Extraer mes y año de FECHA
    df_temp['FECHA_MES'] = df_temp['FECHA'].dt.month
    df_temp['FECHA_AÑO'] = df_temp['FECHA'].dt.year
    df_temp['FECHA_DIA'] = df_temp['FECHA'].dt.day
    
    # Filtrar por mes exacto
    df_mes = df_temp[df_temp['FECHA_MES'] == mes_num].copy()
    print(f"Registros del mes {mes_seleccionado}: {len(df_mes)}")
    
    if df_mes.empty:
        print("DataFrame vacío después de filtrar por mes")
        return pd.DataFrame()
    
    # Si hay múltiples años, tomar el más reciente
    año_filtro = df_mes['FECHA_AÑO'].max()
    print(f"Año filtro: {año_filtro}")
    
    df_mes = df_mes[df_mes['FECHA_AÑO'] == año_filtro]
    print(f"Registros para {mes_seleccionado} {año_filtro}: {len(df_mes)}")
    
    # Filtrar VENTAS - todos los registros sin importar PAGO o ESTADO
    df_ventas = df_mes.copy()
    
    if df_ventas.empty:
        print("DataFrame vacío después de filtrar por año")
        return pd.DataFrame()
    
    # Validar días válidos del mes
    if mes_num == 12:
        último_día_mes = pd.Timestamp(year=año_filtro+1, month=1, day=1) - pd.DateOffset(days=1)
    else:
        último_día_mes = pd.Timestamp(year=año_filtro, month=mes_num+1, day=1) - pd.DateOffset(days=1)
    
    último_día_válido = último_día_mes.day
    print(f"Último día válido del mes: {último_día_válido}")
    
    # Filtrar días válidos
    df_ventas = df_ventas[(df_ventas['FECHA_DIA'] >= 1) & (df_ventas['FECHA_DIA'] <= último_día_válido)]
    print(f"Registros con días válidos: {len(df_ventas)}")
    
    if df_ventas.empty:
        print("DataFrame vacío después de filtrar días")
        return pd.DataFrame()
    
    # Contar por día
    df_dias = df_ventas.groupby('FECHA_DIA').size().reset_index(name='INSTALADAS')
    print(f"Días únicos con datos: {len(df_dias)}")
    print(df_dias)
    
    return df_dias

result = get_instaladas_por_semana_debug("Abril")
print(f"\nResultado final: {len(result)} registros")
