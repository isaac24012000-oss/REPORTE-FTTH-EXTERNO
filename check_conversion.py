import pandas as pd

# Cargar datos
df_mantra = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='MANTRA')
df_drive = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='DRIVE')

mes_actual = 'Junio'

# Con Cobertura en MANTRA
df_mes_mantra = df_mantra[df_mantra['Mes'] == mes_actual].copy()
df_mes_mantra['NIVEL 2'] = df_mes_mantra['NIVEL 2'].astype(str).str.strip()
con_cobertura = len(df_mes_mantra[df_mes_mantra['NIVEL 2'] == 'Con Cobertura'])

print(f"Mes: {mes_actual}")
print(f"Con Cobertura en MANTRA: {con_cobertura}")

# Instaladas en DRIVE
df_drive['FECHA'] = pd.to_datetime(df_drive['FECHA'], errors='coerce')
df_drive_junio = df_drive[(df_drive['FECHA'].dt.month == 6) & (df_drive['FECHA'].dt.year == 2026)].copy()

df_drive_junio['PAGO'] = df_drive_junio['PAGO'].astype(str).str.strip()
df_drive_junio['ESTADO'] = df_drive_junio['ESTADO'].astype(str).str.strip()

# Mostrar todos los estados disponibles
print(f"\nEstados en DRIVE Junio: {df_drive_junio['ESTADO'].unique()}")
print(f"Instancia de PAGO valores: {df_drive_junio['PAGO'].unique()[:10]}")

instaladas_con_pago = len(df_drive_junio[
    (df_drive_junio['ESTADO'] == 'INSTALADO') &
    (df_drive_junio['PAGO'] != '') & 
    (df_drive_junio['PAGO'] != 'nan') & 
    (df_drive_junio['PAGO'].notna())
])

print(f"\nInstaladas con PAGO en DRIVE (Junio): {instaladas_con_pago}")

# Contar solo INSTALADO
solo_instalado = len(df_drive_junio[df_drive_junio['ESTADO'] == 'INSTALADO'])
print(f"Solo INSTALADO: {solo_instalado}")

# Desglose de PAGO
print(f"\nDesglose de PAGO en INSTALADO:")
print(f"  SI: {len(df_drive_junio[(df_drive_junio['ESTADO'] == 'INSTALADO') & (df_drive_junio['PAGO'] == 'SI')])}")
print(f"  NO: {len(df_drive_junio[(df_drive_junio['ESTADO'] == 'INSTALADO') & (df_drive_junio['PAGO'] == 'NO')])}")
print(f"  Vacío/nan: {len(df_drive_junio[(df_drive_junio['ESTADO'] == 'INSTALADO') & ((df_drive_junio['PAGO'] == '') | (df_drive_junio['PAGO'] == 'nan') | (df_drive_junio['PAGO'].isna()))])}")

if con_cobertura > 0:
    conversion = round((instaladas_con_pago / con_cobertura * 100))
    print(f"\nConversión actual (solo con PAGO != ''): {instaladas_con_pago}/{con_cobertura} = {conversion}%")
    
    # Si debería ser 42/28
    if solo_instalado == 42:
        print(f"\n✓ Hay {solo_instalado} INSTALADOS en total")
        conversion_sin_filtro = round((solo_instalado / con_cobertura * 100))
        print(f"Si contar todos INSTALADO (sin filtro PAGO): {solo_instalado}/{con_cobertura} = {conversion_sin_filtro}%")
    
    if instaladas_con_pago != 42:
        print(f"\n⚠️  Debería ser 42, pero obtenemos: {instaladas_con_pago}")
        print(f"Diferencia: {42 - instaladas_con_pago}")
else:
    print("Sin datos de Con Cobertura")
