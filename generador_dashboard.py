import os
import json
import math
import pandas as pd
import datetime

# 1. Localizar archivo Excel en data/
data_dir = 'data'
xlsx_files = [f for f in os.listdir(data_dir) if f.endswith('.xlsx') and not f.startswith('~$')] if os.path.exists(data_dir) else []
excel_path = os.path.join(data_dir, xlsx_files[0]) if xlsx_files else 'Horometros.xlsx'

df_raw = pd.read_excel(excel_path, sheet_name='Hoja 1')

# 2. Detectar dinámicamente las columnas mensuales de fechas
date_cols_dict = {}
for c in df_raw.columns:
    if c in ['FECHA CONTRATO (Resumen)', 'FECHA ACTA INICIO (Resumen)', 'FECHA FIN (Resumen)']:
        continue
    if isinstance(c, (datetime.datetime, pd.Timestamp)):
        date_cols_dict[pd.to_datetime(c)] = c
    else:
        try:
            dt = pd.to_datetime(c)
            if 2020 <= dt.year <= 2040:
                date_cols_dict[dt] = c
        except:
            pass

sorted_dates = sorted(date_cols_dict.keys())
corte_dt = sorted_dates[-1]
col_corte = date_cols_dict[corte_dt]
col_inicio_ano = date_cols_dict[sorted_dates[0]]

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

    val_corte = safe_float(row.get(col_corte), 0.0)

    if pd.isna(f_ini):
        val_ini_ano = safe_float(row.get(col_inicio_ano), 0.0)
        consumo = max(val_corte - val_ini_ano, 0.0)
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
            'meses_ciclo': len(sorted_dates) - 1,
            'ano_texto': "Sin vigencia",
            'estado': "Normal" if consumo <= horas_pactadas_num else "Excedido",
            'regional': regional,
            'historial': [round(safe_float(row.get(date_cols_dict[d])), 1) for d in sorted_dates]
        })
        continue

    expired = False
    if pd.notnull(f_fin) and corte_dt > f_fin and f_fin.year > 2000:
        expired = True
        
    years_passed = corte_dt.year - f_ini.year
    if (corte_dt.month, corte_dt.day) < (f_ini.month, f_ini.day):
        years_passed -= 1
    current_year_num = years_passed + 1
    
    tot_y_str = int(total_years) if total_years > 0 else '?'
    ano_texto = f"Año {current_year_num} de {tot_y_str}"
    if expired:
        ano_texto += " (Vencido)"
        
    if (corte_dt.month, corte_dt.day) >= (f_ini.month, f_ini.day):
        start_cycle_dt = pd.Timestamp(year=corte_dt.year, month=f_ini.month, day=1)
    else:
        start_cycle_dt = pd.Timestamp(year=corte_dt.year - 1, month=f_ini.month, day=1)
        
    fechas_disponibles_ciclo = [d for d in sorted_dates if d >= start_cycle_dt]
    fecha_inicio_evaluada = fechas_disponibles_ciclo[0] if fechas_disponibles_ciclo else sorted_dates[0]

    val_base = safe_float(row.get(date_cols_dict[fecha_inicio_evaluada]), 0.0)
    consumo_periodo = max(val_corte - val_base, 0.0)

    meses_ciclo = (corte_dt.year - fecha_inicio_evaluada.year) * 12 + (corte_dt.month - fecha_inicio_evaluada.month)
    if meses_ciclo <= 0:
        meses_ciclo = 1

    limite_proporcional = (horas_pactadas_num / 12.0) * meses_ciclo
    
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
        'historial': [round(safe_float(row.get(date_cols_dict[d])), 1) for d in sorted_dates]
    })

os.makedirs('public', exist_ok=True)
with open(os.path.join('public', 'dashboard_data.json'), 'w', encoding='utf-8') as f:
    json.dump(clean_data, f, ensure_ascii=False, indent=4)

# 3. Guardar metadatos de sincronización
metadata = {
    "corte_mes": corte_dt.strftime('%B %Y').capitalize(),
    "fecha_corte": corte_dt.strftime('%Y-%m-%d'),
    "ultima_actualizacion": datetime.datetime.now().strftime('%d/%m/%Y %I:%M %p')
}
with open(os.path.join('public', 'metadata.json'), 'w', encoding='utf-8') as f:
    json.dump(metadata, f, ensure_ascii=False, indent=4)

print(f"Completado: Corte {corte_dt.strftime('%Y-%m-%d')} procesado.")
