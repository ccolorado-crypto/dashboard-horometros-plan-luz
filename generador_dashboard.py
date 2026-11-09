import os
import json
import math
import pandas as pd
import datetime

# 1. Localizar archivo Excel en data/ o raíz
data_dir = 'data'
xlsx_files = []
if os.path.exists(data_dir):
    xlsx_files = [
        os.path.join(data_dir, f) for f in os.listdir(data_dir) 
        if f.endswith('.xlsx') and not f.startswith('~$')
    ]

# Si hay varios archivos, ordenar por fecha de modificación (el más reciente primero)
if xlsx_files:
    xlsx_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    excel_path = xlsx_files[0]
else:
    # Buscar en directorio actual
    local_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~$') and 'Horometros' in f]
    if local_files:
        local_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        excel_path = local_files[0]
    else:
        excel_path = 'Horometros.xlsx'

print(f"Leyendo archivo Excel: {excel_path}")
df_raw = pd.read_excel(excel_path, sheet_name='Hoja 1')

# 2. Detectar dinámicamente columnas mensuales
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

meses_es = {
    1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun',
    7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'
}
meses_completos = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
    7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}

mes_corte_str = f"{meses_completos[corte_dt.month]} {corte_dt.year}"
meses_labels = [meses_es[d.month] for d in sorted_dates]

# Fecha y hora de actualización (formato legible en Colombia / UTC-5)
hora_actual = (datetime.datetime.utcnow() - datetime.timedelta(hours=5)).strftime('%d/%m/%Y %I:%M %p')

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
    if fechas_disponibles_ciclo:
        fecha_inicio_evaluada = fechas_disponibles_ciclo[0]
    else:
        fecha_inicio_evaluada = sorted_dates[0]

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

payload = {
    'corte_mes': mes_corte_str,
    'corte_fecha': corte_dt.strftime('%Y-%m-%d'),
    'ultima_actualizacion': hora_actual,
    'meses_labels': meses_labels,
    'equipos': clean_data
}

os.makedirs('public', exist_ok=True)
with open(os.path.join('public', 'dashboard_data.json'), 'w', encoding='utf-8') as f:
    json.dump(payload, f, ensure_ascii=False, indent=4)

print(f"Generado exitosamente: {len(clean_data)} generadores procesados al corte {mes_corte_str} ({hora_actual}).")
