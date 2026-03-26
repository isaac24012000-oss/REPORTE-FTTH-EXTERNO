import pandas as pd
import warnings
warnings.filterwarnings('ignore')

# Cargar datos
df_drive = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='DRIVE')
df_mantra = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='MANTRA')

# Filtrar marzo
df_drive_marzo = df_drive[df_drive['MES'] == 'Marzo'].copy()
df_mantra_marzo = df_mantra[df_mantra['Mes'] == 'Marzo'].copy()

asesor = 'ZIM_PERCYMC_VTP'

print(f'=== CONVERSIÓN PARA {asesor} ===\n')

# Transacciones con PAGO en DRIVE
df_drive_temp = df_drive_marzo[df_drive_marzo['ASESOR'] == asesor].copy()
print(f'Registros DRIVE para {asesor}: {len(df_drive_temp)}')
print(f'Estados en DRIVE: {dict(df_drive_temp["ESTADO"].value_counts())}')

df_drive_temp.loc[:, 'PAGO'] = df_drive_temp['PAGO'].astype(str).str.strip()
pago_count = len(df_drive_temp[(df_drive_temp['PAGO'] != '') & (df_drive_temp['PAGO'] != 'nan')])
print(f'Transacciones con PAGO: {pago_count}')

# Con Cobertura en MANTRA
df_mantra_temp = df_mantra_marzo[df_mantra_marzo['Agente'].str.strip() == asesor].copy()
print(f'\nRegistros MANTRA para {asesor}: {len(df_mantra_temp)}')
if len(df_mantra_temp) > 0:
    df_mantra_temp.loc[:, 'NIVEL 2'] = df_mantra_temp['NIVEL 2'].astype(str).str.strip()
    nivel2_counts = dict(df_mantra_temp['NIVEL 2'].value_counts())
    print(f'Distribución NIVEL 2: {nivel2_counts}')
    cobertura_count = len(df_mantra_temp[df_mantra_temp['NIVEL 2'] == 'Con Cobertura'])
    print(f'Con Cobertura: {cobertura_count}')
else:
    cobertura_count = 0
    print('No hay registros en MANTRA para este asesor en Marzo')

conversion = round((pago_count / cobertura_count * 100)) if cobertura_count > 0 else 0
print(f'\nConversión: ({pago_count} / {cobertura_count}) * 100 = {conversion}%')
