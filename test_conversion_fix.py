import pandas as pd

# Cargar datos
df_mantra = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='MANTRA')
df_drive = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='DRIVE')

mes_actual = 'Junio'

# Con Cobertura en MANTRA
df_mes_mantra = df_mantra[df_mantra['Mes'] == mes_actual].copy()
df_mes_mantra['NIVEL 2'] = df_mes_mantra['NIVEL 2'].astype(str).str.strip()
con_cobertura = len(df_mes_mantra[df_mes_mantra['NIVEL 2'] == 'Con Cobertura'])

print(f"Con Cobertura: {con_cobertura}")

# Instaladas + Pendiente en DRIVE
df_drive['FECHA'] = pd.to_datetime(df_drive['FECHA'], errors='coerce')

# Mes numero
mes_numeros = {
    'Enero': 1, 'Febrero': 2, 'Marzo': 3, 'Abril': 4,
    'Mayo': 5, 'Junio': 6, 'Julio': 7, 'Agosto': 8,
    'Septiembre': 9, 'Octubre': 10, 'Noviembre': 11, 'Diciembre': 12
}
mes_num = mes_numeros.get(mes_actual, None)

# Determinar año
año_actual = df_drive['FECHA'].dt.year.max()

# Crear fecha de inicio (1/mes/año)
fecha_inicio = pd.Timestamp(year=año_actual, month=mes_num, day=1)

# Crear fecha de fin (último día del mes)
if mes_num == 12:
    fecha_fin = pd.Timestamp(year=año_actual + 1, month=1, day=1) - pd.Timedelta(days=1)
else:
    fecha_fin = pd.Timestamp(year=año_actual, month=mes_num + 1, day=1) - pd.Timedelta(days=1)

df_mes_drive = df_drive[(df_drive['FECHA'] >= fecha_inicio) & (df_drive['FECHA'] <= fecha_fin)].copy()

# Contar INSTALADAS + PENDIENTE
df_mes_drive['ESTADO'] = df_mes_drive['ESTADO'].astype(str).str.strip()

total_transacciones = len(df_mes_drive[
    (df_mes_drive['ESTADO'] == 'INSTALADO') |
    (df_mes_drive['ESTADO'] == 'PENDIENTE')
])

print(f"Total transacciones (INSTALADO + PENDIENTE): {total_transacciones}")

if total_transacciones > 0:
    conversion_pct = round((total_transacciones / con_cobertura * 100)) if con_cobertura > 0 else 0
    print(f"Conversion: {total_transacciones}/{con_cobertura} = {conversion_pct}%")
    
    if total_transacciones == 42 and con_cobertura == 28:
        print(f"\n✓ Correcto: 42/28 = {conversion_pct}%")
