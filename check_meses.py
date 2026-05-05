import pandas as pd

def get_meses_disponibles():
    """Obtiene dinámicamente todos los meses disponibles en los datos
    Retorna lista de tuplas (mes_año, mes_nombre, año, mes_num)"""
    df_mantra = pd.read_excel('REPORTE FTTH.xlsx', sheet_name='MANTRA')
    
    if df_mantra is None or df_mantra.empty:
        # Fallback a los meses por defecto
        return [
            ('Noviembre 2025', 'Noviembre', 2025, 11),
            ('Diciembre 2025', 'Diciembre', 2025, 12),
            ('Enero 2026', 'Enero', 2026, 1),
            ('Febrero 2026', 'Febrero', 2026, 2)
        ]
    
    # Obtener meses únicos
    meses_unicos = df_mantra['Mes'].dropna().unique().tolist()
    
    if not meses_unicos:
        return [
            ('Noviembre 2025', 'Noviembre', 2025, 11),
            ('Diciembre 2025', 'Diciembre', 2025, 12),
            ('Enero 2026', 'Enero', 2026, 1),
            ('Febrero 2026', 'Febrero', 2026, 2)
        ]
    
    # Orden correcto de meses
    orden_meses = [
        'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
        'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
    ]
    
    # Mapeo de meses a números
    mes_num_map = {
        'Enero': 1, 'Febrero': 2, 'Marzo': 3, 'Abril': 4,
        'Mayo': 5, 'Junio': 6, 'Julio': 7, 'Agosto': 8,
        'Septiembre': 9, 'Octubre': 10, 'Noviembre': 11, 'Diciembre': 12
    }
    
    # Filtrar y ordenar meses según el orden correcto
    meses_ordenados = [m for m in orden_meses if m in meses_unicos]
    
    # Si hay meses no en el orden (edge case), agregarlos al final
    meses_faltantes = [m for m in meses_unicos if m not in orden_meses]
    meses_ordenados.extend(meses_faltantes)
    
    # Crear tuplas (mes_año, mes_nombre, año, mes_num)
    # Asumir año basado en si es antes o después de junio
    resultado = []
    for mes_nombre in meses_ordenados:
        mes_num = mes_num_map.get(mes_nombre, 1)
        # Si es Noviembre-Diciembre, asumir 2025; si es Enero-Octubre, asumir 2026
        año = 2025 if mes_num >= 11 else 2026
        mes_año = f"{mes_nombre} {año}"
        resultado.append((mes_año, mes_nombre, año, mes_num))
    
    return resultado

meses_disponibles = get_meses_disponibles()
print("Meses disponibles:")
for mes_año, mes_nombre, año, mes_num in meses_disponibles:
    print(f"  {mes_año}")
    
print("\nOpciones para selectbox:")
opciones_meses = [mes_año for mes_año, _, _, _ in meses_disponibles]
print(opciones_meses)
