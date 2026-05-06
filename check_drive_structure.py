import pandas as pd

# Verificar estructura de DRIVE
df_drive = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='DRIVE')
print("Columnas en DRIVE:")
print(df_drive.columns.tolist())
print("\nPrimeras filas:")
print(df_drive[['FECHA', 'MES']].head(10))
print("\nMeses en DRIVE:")
print(df_drive['MES'].dropna().unique())
print("\nFechas en DRIVE:")
print(f"Mínima: {df_drive['FECHA'].min()}")
print(f"Máxima: {df_drive['FECHA'].max()}")

# Detectar mes más reciente
df_drive['FECHA'] = pd.to_datetime(df_drive['FECHA'], errors='coerce')
df_drive_sin_nulos = df_drive[df_drive['FECHA'].notna()].copy()
df_drive_sin_nulos['FECHA_MES'] = df_drive_sin_nulos['FECHA'].dt.month
df_drive_sin_nulos['FECHA_AÑO'] = df_drive_sin_nulos['FECHA'].dt.year

# Agrupar por mes y año, y encontrar el más reciente
mes_año_max = df_drive_sin_nulos.groupby(['FECHA_AÑO', 'FECHA_MES']).size().reset_index()
print("\nMeses únicos en DRIVE (AÑO-MES):")
print(mes_año_max.sort_values(['FECHA_AÑO', 'FECHA_MES'], ascending=False))
