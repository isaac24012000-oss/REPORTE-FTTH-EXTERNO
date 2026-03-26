import pandas as pd
import numpy as np

# Cargar datos
df_drive = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='DRIVE')
df_mantra = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='MANTRA')

# Verificar datos de Marzo
df_marzo_drive = df_drive[df_drive['MES'] == 'Marzo']
df_marzo_mantra = df_mantra[df_mantra['Mes'] == 'Marzo']

print("=== DRIVE - Marzo ===")
print(f"Total registros: {len(df_marzo_drive)}")
print(f"Valores nulos por columna:")
print(df_marzo_drive.isnull().sum())
print(f"\nAsesores únicos: {df_marzo_drive['ASESOR'].nunique()}")
print(f"Primeros asesores: {df_marzo_drive['ASESOR'].unique()[:5].tolist()}")

print("\n=== MANTRA - Marzo ===")
print(f"Total registros: {len(df_marzo_mantra)}")
print(f"Valores nulos por columna:")
print(df_marzo_mantra.isnull().sum())

print("\n✅ Todos los datos están listos para cargar en Streamlit")
