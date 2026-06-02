"""
generar_dashboard.py
--------------------
Lee el Excel de contramuestras, procesa AM y DM,
y genera un archivo HTML con el dashboard actualizado.

Uso:
    python generar_dashboard.py

Requisitos:
    pip install pandas openpyxl
"""

import pandas as pd
import json
import re
import sys
from pathlib import Path

# ─── CONFIGURACIÓN ──────────────────────────────────────────────────────────
EXCEL_FILE   = "Registro_De_Contramuestras_Palta_2026.xlsm"   # <── cambia si el nombre cambia
OUTPUT_HTML  = "dashboard_contramuestras.html"
SEMANA_LABEL = "Semana 23"   # <── actualiza cada semana
FECHA_LABEL  = "01–06–2026"  # <── actualiza cada semana
# ────────────────────────────────────────────────────────────────────────────

def read_sheet(xl_path, sheet, key_col):
    """Lee una hoja buscando la fila de encabezado."""
    df_raw = pd.read_excel(xl_path, sheet_name=sheet, header=None)
    hr = 0
    for i, row in df_raw.iterrows():
        if str(row[0]).strip() == key_col:
            hr = i; break
    return pd.read_excel(xl_path, sheet_name=sheet, header=hr)

def clean_cols(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df

def norm_destino(s):
    return {'EE. UU.':'USA','Reino Unido':'UK'}.get(s, s)

# ── Leer AM ─────────────────────────────────────────────────────────────────
print("Leyendo hoja AM...")
df_am = clean_cols(read_sheet(EXCEL_FILE, 'AM', 'N° de Muestra'))
df_am = df_am[df_am['Días De Evalución'] != '0 Días'].dropna(subset=['Fundo','Días De Evalución'])
df_am['País Destino'] = df_am['País Destino'].apply(norm_destino)
df_am['Calibre'] = df_am['Calibre'].astype(float).astype(int)

AM_METRICS = ['Nivel Daño Lenticelar','Nivel Moho Peduncular',
              'Nivel Mancha Por Senecensia','Nivel Manchas Por Frio',
              'Nivel Pulpa Gris','Nivel Haces Vasculares']
AM_GRP = ['Semana','Fundo','Días De Evalución','País Destino','Calibre','Fungicida']

am_records = []
for name, g in df_am.groupby(AM_GRP):
    r = dict(zip(AM_GRP, name))
    r['Días'] = r.pop('Días De Evalución')
    r['n'] = len(g)
    r['dl']  = {int(k): round(int(v)/len(g)*100,1) for k,v in g['Nivel Daño Lenticelar'].value_counts().items()}
    r['mp']  = {int(k): round(int(v)/len(g)*100,1) for k,v in g['Nivel Moho Peduncular'].value_counts().items()}
    r['ms']  = {int(k): round(int(v)/len(g)*100,1) for k,v in g['Nivel Mancha Por Senecensia'].value_counts().items()} if 'Nivel Mancha Por Senecensia' in g else {'0':100}
    r['mf']  = {int(k): round(int(v)/len(g)*100,1) for k,v in g['Nivel Manchas Por Frio'].value_counts().items()} if 'Nivel Manchas Por Frio' in g else {'0':100}
    r['pg']  = {int(k): round(int(v)/len(g)*100,1) for k,v in g['Nivel Pulpa Gris'].value_counts().items()} if 'Nivel Pulpa Gris' in g else {'0':100}
    r['hv']  = {int(k): round(int(v)/len(g)*100,1) for k,v in g['Nivel Haces Vasculares'].value_counts().items()} if 'Nivel Haces Vasculares' in g else {'0':100}
    am_records.append(r)

print(f"  AM: {len(am_records)} grupos procesados")

# ── Leer DM ─────────────────────────────────────────────────────────────────
print("Leyendo hoja DM...")
df_dm = clean_cols(read_sheet(EXCEL_FILE, 'DM', 'N° De Muestra'))
df_dm = df_dm.dropna(subset=['Fundo','Días De Evaluación'])
df_dm['País Destino'] = df_dm['País Destino'].apply(norm_destino)
df_dm['Calibre'] = df_dm['Calibre'].astype(float).astype(int)

DM_MK_MAP = {
    'Nivel De Pulpa Gris':             'pg',
    'Nivel Pudrición Pulpa':           'pp',
    'Nivel De Pudrición Peduncular':   'ppu',
    'Nivel De  Hongo Peduncular':      'hp',
    'Nivel De Maduración Hetérogenea': 'mh',
    'Nivel De Manchas Externas':       'me',
}
DM_GRP = ['Semana','Fundo','Días De Evaluación','País Destino','Calibre','Fungicida']

dm_records = []
for name, g in df_dm.groupby(DM_GRP):
    r = dict(zip(DM_GRP, name))
    r['Días'] = r.pop('Días De Evaluación')
    r['n'] = len(g)
    fw_col = 'Promedio Firmeza' if 'Promedio Firmeza' in g.columns else None
    dm_col = 'Días De Maduración' if 'Días De Maduración' in g.columns else None
    r['fw'] = round(float(g[fw_col].mean()),2) if fw_col and g[fw_col].notna().any() else 0
    r['dm'] = round(float(g[dm_col].mean()),1) if dm_col and g[dm_col].notna().any() and g[dm_col].mean() > 0 and g[dm_col].mean() < 100 else 0
    for orig, short in DM_MK_MAP.items():
        col = next((c for c in g.columns if c.strip() == orig.strip()), None)
        if col:
            r[short] = {int(k): round(int(v)/len(g)*100,1) for k,v in g[col].value_counts().items()}
        else:
            r[short] = {0: 100.0}
    dm_records.append(r)

print(f"  DM: {len(dm_records)} grupos procesados")

# ── Serializar ───────────────────────────────────────────────────────────────
am_json = json.dumps(am_records, ensure_ascii=False, separators=(',',':'))
dm_json = json.dumps(dm_records, ensure_ascii=False, separators=(',',':'))

print(f"  Tamaño datos: AM={len(am_json)//1024}KB, DM={len(dm_json)//1024}KB")

# ── Leer template HTML e inyectar datos ──────────────────────────────────────
TEMPLATE = Path("dashboard_template.html")
if not TEMPLATE.exists():
    print(f"\n❌ No se encontró '{TEMPLATE}'.")
    print("   Asegúrate de tener el archivo template en la misma carpeta.")
    sys.exit(1)

html = TEMPLATE.read_text(encoding='utf-8')

# Reemplazar los arrays de datos
html = re.sub(r'const AM=\[.*?\];', f'const AM={am_json};', html, flags=re.DOTALL)
html = re.sub(r'const DM=\[.*?\];', f'const DM={dm_json};', html, flags=re.DOTALL)

# Reemplazar semana y fecha en el header
html = re.sub(r'Semana \d+', SEMANA_LABEL, html)
html = re.sub(r'\d{2}–\d{2}–\d{4}', FECHA_LABEL, html)

Path(OUTPUT_HTML).write_text(html, encoding='utf-8')
print(f"\n✅ Dashboard generado: {OUTPUT_HTML}")
print(f"   Semanas disponibles en AM: {sorted(df_am['Semana'].unique().tolist())}")
print(f"   Semanas disponibles en DM: {sorted(df_dm['Semana'].unique().tolist())}")
