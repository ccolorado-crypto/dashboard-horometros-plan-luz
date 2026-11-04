import os
import json
import pandas as pd

data_dir = 'data'
xlsx_files = [f for f in os.listdir(data_dir) if f.endswith('.xlsx')] if os.path.exists(data_dir) else []
excel_path = os.path.join(data_dir, xlsx_files[0]) if xlsx_files else 'Horometros .xlsx'

df_raw = pd.read_excel(excel_path, sheet_name='Hoja 1')

hist_cols = [c for c in df_raw.columns if '2026' in str(c)]
jan_col = hist_cols[0]
sep_col = hist_cols[-1]

df_raw['Consumo_2026'] = df_raw[sep_col] - df_raw[jan_col]

clean_data = []
for index, row in df_raw.iterrows():
    fecha_inicio = row.get('FECHA ACTA INICIO (Resumen)')
    fecha_fin = row.get('FECHA FIN (Resumen)')
    
    inicio_str = fecha_inicio.strftime('%Y-%m-%d') if pd.notnull(fecha_inicio) and hasattr(fecha_inicio, 'strftime') else "N/A"
    fin_str = fecha_fin.strftime('%Y-%m-%d') if pd.notnull(fecha_fin) and hasattr(fecha_fin, 'strftime') else "N/A"
    
    try:
        horas_pactadas_num = float(row.get('Horas x contrato', 200))
        if pd.isna(horas_pactadas_num):
            horas_pactadas_num = 200.0
    except:
        horas_pactadas_num = 200.0
        
    consumo = float(row['Consumo_2026']) if pd.notnull(row['Consumo_2026']) else 0.0
    potencia = float(row.get('POTENCIA kW', 0)) if pd.notnull(row.get('POTENCIA kW')) else 0.0
    anos = float(row.get('AÑOS CONTRATO (Resumen)', 0)) if pd.notnull(row.get('AÑOS CONTRATO (Resumen)')) else 0.0
    
    limite_proporcional = (horas_pactadas_num / 12.0) * 8.0
    if consumo == 0:
        estado = "Inactivo"
    elif consumo > horas_pactadas_num:
        estado = "Excedido"
    elif consumo > limite_proporcional:
        estado = "Alerta"
    else:
        estado = "Normal"
        
    id_val = str(row.get('ID', 'N/A')).strip()
    regional = id_val[:3] if len(id_val) >= 3 else "OTRO"
    if regional not in ['MED', 'BOG', 'BQL', 'CAL', 'PEI', 'IBG']:
        regional = "OTRO"

    item = {
        "id": id_val,
        "nombre": str(row.get('Nombre Artimo', 'N/A')),
        "potencia": potencia,
        "horas_pactadas": horas_pactadas_num,
        "contrato": str(row.get('N° CONTRATO', 'N/A')),
        "anos_contrato": anos,
        "fecha_inicio": inicio_str,
        "fecha_fin": fin_str,
        "consumo_2026": round(consumo, 1),
        "estado": estado,
        "regional": regional,
        "historial": [round(float(row[c]), 1) if pd.notnull(row[c]) else 0.0 for c in hist_cols]
    }
    clean_data.append(item)

os.makedirs('public', exist_ok=True)
with open(os.path.join('public', 'dashboard_data.json'), 'w', encoding='utf-8') as f:
    json.dump(clean_data, f, ensure_ascii=False, indent=4)

print(f"Generado exitosamente con {len(clean_data)} generadores.")
