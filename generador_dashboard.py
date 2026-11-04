import os
import json
import math
import pandas as pd
import datetime

data_dir = 'data'
xlsx_files = [f for f in os.listdir(data_dir) if f.endswith('.xlsx')] if os.path.exists(data_dir) else []
excel_path = os.path.join(data_dir, xlsx_files[0]) if xlsx_files else 'Horometros .xlsx'

df_raw = pd.read_excel(excel_path, sheet_name='Hoja 1')
corte_dt = pd.Timestamp('2026-09-01')

# Detect monthly 2026 columns
hist_cols = [c for c in df_raw.columns if '2026' in str(c)]
date_cols_dict = {}
for c in hist_cols:
    if isinstance(c, datetime.datetime):
        date_cols_dict[c.month] = c
    else:
        try:
            parsed = pd.to_datetime(c)
            date_cols_dict[parsed.month] = c
        except:
            pass

def safe_float(val, default=0.0):
    try:
        if pd.isna(val) or val is None:
            return default
        f = float(val)
        if math.isnan(f) or math.isinf(f):
            return default
        return f
    except:
        return default

clean_data = []
for index, row in df_raw.iterrows():
    f_ini = row.get('FECHA ACTA INICIO (Resumen)')
    f_fin = row.get('FECHA FIN (Resumen)')
    
    inicio_str = f_ini.strftime('%Y-%m-%d') if pd.notnull(f_ini) and hasattr(f_ini, 'strftime') else "N/A"
    fin_str = f_fin.strftime('%Y-%m-%d') if pd.notnull(f_fin) and hasattr(f_fin, 'strftime') else "N/A"

    horas_pactadas_num = safe_float(row.get('Horas x contrato'), 200.0)
    potencia = safe_float(row.get('POTENCIA kW'), 0.0)
    total_years = safe_float(row.get('AÑOS CONTRATO (Resumen)'), 0.0)
    
    id_val = str(row.get('ID', 'N/A')).strip() if pd.notnull(row.get('ID')) else "N/A"
    regional = id_val[:3] if len(id_val) >= 3 else "OTRO"
    if regional not in ['MED', 'BOG', 'BQL', 'CAL', 'PEI', 'IBG']:
        regional = "OTRO"

    if pd.isna(f_ini):
        # Fallback if no start date
        consumo = safe_float(row.get(hist_cols[-1])) - safe_float(row.get(hist_cols[0]))
        clean_data.append({
            'id': id_val,
            'nombre': str(row.get('Nombre Artimo', 'N/A')),
            'potencia': potencia,
            'horas_pactadas': horas_pactadas_num,
            'contrato': str(row.get('N° CONTRATO', 'N/A')),
            'anos_contrato': total_years,
            'fecha_inicio': inicio_str,
            'fecha_fin': fin_str,
            'consumo_2026': round(consumo, 1),
            'ano_texto': "Sin vigencia",
            'estado': "Normal" if consumo <= horas_pactadas_num else "Excedido",
            'regional': regional,
            'historial': [round(safe_float(row[c], 0.0), 1) for c in hist_cols]
        })
        continue
        
    # Check if expired
    expired = False
    if pd.notnull(f_fin) and corte_dt > f_fin and f_fin.year > 2000:
        expired = True
        
    # Calculate current contract year
    years_passed = corte_dt.year - f_ini.year
    if (corte_dt.month, corte_dt.day) < (f_ini.month, f_ini.day):
        years_passed -= 1
    current_year_num = years_passed + 1
    
    tot_y_str = int(total_years) if total_years > 0 else '?'
    ano_texto = f"Año {current_year_num} de {tot_y_str}"
    if expired:
        ano_texto += " (Vencido)"
        
    start_month = f_ini.month
    col_sep = date_cols_dict.get(9, hist_cols[-1])
    val_sep = safe_float(row.get(col_sep), 0.0)
    
    # Contract Year active logic for 2026 slice
    if 1 <= start_month <= 8:
        col_inicio_ciclo = date_cols_dict.get(start_month, hist_cols[0])
        val_inicio = safe_float(row.get(col_inicio_ciclo), 0.0)
        consumo_periodo = val_sep - val_inicio
        meses_ciclo = 9 - start_month
    else:
        col_inicio_ciclo = date_cols_dict.get(1, hist_cols[0])
        val_inicio = safe_float(row.get(col_inicio_ciclo), 0.0)
        consumo_periodo = val_sep - val_inicio
        meses_ciclo = 8
        
    if consumo_periodo < 0:
        consumo_periodo = 0.0
        
    limite_proporcional = (horas_pactadas_num / 12.0) * max(meses_ciclo, 1)
    
    if consumo_periodo == 0:
        estado = "Inactivo"
    elif consumo_periodo > horas_pactadas_num:
        estado = "Excedido"
    elif consumo_periodo > limite_proporcional:
        estado = "Alerta"
    else:
        estado = "Normal"

    clean_data.append({
        'id': id_val,
        'nombre': str(row.get('Nombre Artimo', 'N/A')),
        'potencia': potencia,
        'horas_pactadas': horas_pactadas_num,
        'contrato': str(row.get('N° CONTRATO', 'N/A')),
        'anos_contrato': total_years,
        'fecha_inicio': inicio_str,
        'fecha_fin': fin_str,
        'consumo_2026': round(consumo_periodo, 1),
        'meses_ciclo': meses_ciclo,
        'ano_texto': ano_texto,
        'estado': estado,
        'regional': regional,
        'historial': [round(safe_float(row[c], 0.0), 1) for c in hist_cols]
    })

os.makedirs('public', exist_ok=True)
with open(os.path.join('public', 'dashboard_data.json'), 'w', encoding='utf-8') as f:
    json.dump(clean_data, f, ensure_ascii=False, indent=4)

print(f"Generado con éxito: {len(clean_data)} generadores procesados (Año en curso).")

