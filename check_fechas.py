import pandas as pd

# Cargar datos
df_mantra = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='MANTRA')
df_drive = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='DRIVE')

mes_actual = 'Junio'

# Con Cobertura en MANTRA
df_mes_mantra = df_mantra[df_mantra['Mes'] == mes_actual].copy()
print(f"Mes: {mes_actual}")
print(f"Rango de fechas en MANTRA Junio: {df_mes_mantra['Fecha'].min()} a {df_mes_mantra['Fecha'].max()}")

# Instaladas en DRIVE
df_drive['FECHA'] = pd.to_datetime(df_drive['FECHA'], errors='coerce')
print(f"\nRango de fechas en DRIVE completo: {df_drive['FECHA'].min()} a {df_drive['FECHA'].max()}")

# Verificar datos de junio 2026
df_drive_junio = df_drive[(df_drive['FECHA'].dt.month == 6) & (df_drive['FECHA'].dt.year == 2026)].copy()
print(f"\nRango de fechas en DRIVE Junio 2026: {df_drive_junio['FECHA'].min()} a {df_drive_junio['FECHA'].max()}")
print(f"Total registros Junio 2026: {len(df_drive_junio)}")

# Instalados en junio
instalados = len(df_drive_junio[df_drive_junio['ESTADO'] == 'INSTALADO'])
print(f"INSTALADOS: {instalados}")

# Mostrar rango completo de fechas INSTALADAS
if instalados > 0:
    instalados_df = df_drive_junio[df_drive_junio['ESTADO'] == 'INSTALADO']
    print(f"Rango de INSTALADOS: {instalados_df['FECHA'].min()} a {instalados_df['FECHA'].max()}")

# Verificar si hay datos de junio en DRIVE (cualquier estado)
print(f"\nDistribución de ESTADOS en Junio:")
print(df_drive_junio['ESTADO'].value_counts())

# Quizás debería contar TODOS los registros (incluyendo PENDIENTE)?
print(f"\nTotal de registros PENDIENTE en Junio: {len(df_drive_junio[df_drive_junio['ESTADO'] == 'PENDIENTE'])}")
print(f"INSTALADO + PENDIENTE = {len(df_drive_junio[(df_drive_junio['ESTADO'] == 'INSTALADO') | (df_drive_junio['ESTADO'] == 'PENDIENTE')])}")
