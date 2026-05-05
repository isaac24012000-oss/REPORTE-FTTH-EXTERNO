import pandas as pd
from datetime import datetime

# Cargar DRIVE
df_drive = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='DRIVE')
df_drive['FECHA'] = pd.to_datetime(df_drive['FECHA'], errors='coerce')

# Verificar fechas de Abril 2026
print("Fechas en el DRIVE:")
print(f"Fecha mínima: {df_drive['FECHA'].min()}")
print(f"Fecha máxima: {df_drive['FECHA'].max()}")
print(f"Hoy es: {datetime.now()}")

# Filtrar por Abril 2026
df_abril = df_drive[df_drive['FECHA'].dt.month == 4].copy()
df_abril = df_abril[df_drive['FECHA'].dt.year == 2026]
print(f"\nRegistros de Abril 2026: {len(df_abril)}")

if len(df_abril) > 0:
    print(f"Primeros registros de Abril:")
    print(df_abril[['FECHA', 'MES']].head())
    print(f"\nMES único en Abril: {df_abril['MES'].unique()}")
