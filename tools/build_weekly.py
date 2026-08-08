"""
build_weekly.py — Dashboard semanal Amazon (sell-out por ASIN)
==============================================================
Genera un HTML interactivo con:
  - Semáforo por semana (≥90% media = verde, 60-89% = ámbar, <60% = rojo)
  - YOY etiquetado (uds / €) según filtro activo
  - Tendencia: W4 vs media W1-W3
  - Tabla 3 niveles colapsable: Familia → Subfamilia → ASIN
  - KPIs + gráficos resumen en cabecera

FUENTE DE DATOS: MerchantSpring (generateOrderedRevenueReport, mercado ES)
CATÁLOGO: Excel con hojas por familia (o colección), con columnas:
  - Hojas normales: A=Subfamilia B=ASIN C=Anualidad D=Descripción
  - Hojas colección: A=Familia B=Subfamilia C=ASIN D=Anualidad E=Descripción

PARA ADAPTAR A OTRO CLIENTE: editar únicamente el bloque ── CONFIG ─────────
"""

import sys, csv, os, json
sys.stdout.reconfigure(encoding='utf-8')

# ═══════════════════════════════════════════════════════════════════
#  ── CONFIG ──  Editar aquí para cada cliente / periodo
# ═══════════════════════════════════════════════════════════════════

# -- Identidad del cliente -------------------------------------------
CLIENT_NAME       = "Finocam"          # nombre visible en el header
CLIENT_LETTER     = "F"                # letra del logo (1 carácter)
MARKETPLACE_LABEL = "ES"               # etiqueta del mercado (badge)
MARKETPLACE_FLAG  = "🇪🇸"              # emoji bandera

# -- Rutas -----------------------------------------------------------
# CSV_DIR: carpeta con los 4 archivos weekly_0.csv … weekly_3.csv
# Estos se descargan desde MerchantSpring (generateOrderedRevenueReport)
# y se renombran como weekly_0, weekly_1, weekly_2, weekly_3 (W1→W4)
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_DIR    = os.environ.get('WEEKLY_CSV_DIR',  os.path.join(_SCRIPT_DIR, 'data'))
EXCEL_PATH = os.environ.get('WEEKLY_EXCEL',    r'C:\Users\Daniela\Downloads\FINOCAM_Familias_Subfamilias_ASIN.xlsx')
CATALOG_JSON = os.path.join(_SCRIPT_DIR, 'data', 'catalog_es.json')
OUT_PATH   = os.environ.get('WEEKLY_OUT_PATH', os.path.join(os.path.dirname(_SCRIPT_DIR), 'FINOCAM_Weekly.html'))

# -- Semanas (W1 → W4, en orden cronológico) -------------------------
# Si WEEKLY_AUTO=1 (modo agente/cloud), se calculan las últimas 4 semanas ISO
import datetime as _dt
_MONTHS_ES = ['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
def _auto_weeks():
    today = _dt.date.today()
    this_mon = today - _dt.timedelta(days=today.weekday())
    out = []
    for i in range(4, 0, -1):
        mon = this_mon - _dt.timedelta(weeks=i)
        sun = mon + _dt.timedelta(days=6)
        label = f'{mon.day}–{sun.day} {_MONTHS_ES[mon.month]}'
        out.append({'label': label, 'wk': f'W{5-i}'})
    return out

WEEKS = _auto_weeks() if os.environ.get('WEEKLY_AUTO') == '1' else [
    {'label': '7–13 Jul',      'wk': 'W1'},
    {'label': '14–20 Jul',     'wk': 'W2'},
    {'label': '21–27 Jul',     'wk': 'W3'},
    {'label': '28 Jul–3 Ago',  'wk': 'W4'},
]

# -- Catálogo: hojas "colección" (cross-familia) ---------------------
# Estas hojas tienen estructura diferente (col A=Familia B=Sub C=ASIN D=Anual E=Desc)
# Dejar vacío: COLS = set()
COLS = {'Moniquilla', 'Talkual'}

# Hojas a ignorar (resumen, índices, etc.)
SKIP = {'Resumen'}

# Validador de anualidad: qué valores se consideran "catálogo activo"
def is_valid_anualidad(val):
    s = str(val or '').strip()
    return s.startswith('2026') or s.startswith('2027') or s == 'NOCAD'

# -- Orden de categorías en la tabla ---------------------------------
# Las familias que no estén aquí van al bloque "Resto" al final
# '_RESTO_' es un comodín obligatorio: captura las familias no listadas
CAT_ORDER = [
    'Moniquilla', 'Talkual',
    'Agendas', 'Calendarios', 'Cuadernos',
    'Planificadores', 'Libros de Firma', 'Índices',
    'Portadocumentos', 'Recambios',
    '_RESTO_',
]

# -- Fecha de actualización (se muestra en el toolbar) ---------------
UPDATE_DATE = _dt.date.today().isoformat() if os.environ.get('WEEKLY_AUTO') == '1' else "2026-08-07"

# ═══════════════════════════════════════════════════════════════════
#  FIN CONFIG — el resto no necesita editarse para un cliente nuevo
# ═══════════════════════════════════════════════════════════════════

# ── 1. Catálogo ────────────────────────────────────────────────────
# Lee desde JSON (cloud/auto) si está disponible; si no, desde Excel (local)
if os.path.exists(CATALOG_JSON):
    with open(CATALOG_JSON, encoding='utf-8') as f:
        catalog = json.load(f)
    print(f'Catálogo: {len(catalog)} ASINs (desde catalog_es.json)')
else:
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    catalog = {}
    for sname in wb.sheetnames:
        if sname in SKIP or sname in COLS:
            continue
        for r in list(wb[sname].iter_rows(values_only=True))[1:]:
            if not r[1]:
                continue
            asin = str(r[1]).strip()
            if not is_valid_anualidad(r[2]):
                continue
            if asin not in catalog:
                catalog[asin] = {
                    'familia': sname,
                    'sub':     str(r[0] or '').strip(),
                    'any':     str(r[2] or '').strip(),
                    'desc':    str(r[3] or '').strip(),
                    'col':     None,
                }
    for col in COLS:
        if col not in wb.sheetnames:
            continue
        for r in list(wb[col].iter_rows(values_only=True))[1:]:
            if not r[2]:
                continue
            asin = str(r[2]).strip()
            if not is_valid_anualidad(r[3]):
                continue
            if asin in catalog:
                catalog[asin]['col'] = col
            else:
                catalog[asin] = {
                    'familia': str(r[0] or '').strip(),
                    'sub':     str(r[1] or '').strip(),
                    'any':     str(r[3] or '').strip(),
                    'desc':    str(r[4] or '').strip(),
                    'col':     col,
                }
    print(f'Catálogo: {len(catalog)} ASINs (desde Excel)')

# ── 2. CSVs semanales ──────────────────────────────────────────────
# Esperados: weekly_0.csv (W1) … weekly_3.csv (W4) en CSV_DIR
# Columnas MerchantSpring: asin, orderedRevenue, orderedUnits,
#                          priorOrderedRevenue, priorOrderedUnits, title
aw = {}     # {asin: {week_index: {rev, u, pr, pu}}}
titles = {}
for wi in range(4):
    csv_path = os.path.join(CSV_DIR, f'weekly_{wi}.csv')
    if not os.path.exists(csv_path):
        print(f'  AVISO: no encontrado {csv_path}')
        continue
    with open(csv_path, encoding='utf-8-sig') as f:
        for row in csv.DictReader(f):
            asin = row['asin']
            if asin not in catalog:
                continue
            try:
                rev = float(row['orderedRevenue'])
                u   = int(row['orderedUnits'])
                pr  = float(row['priorOrderedRevenue'])
                pu  = int(row['priorOrderedUnits'])
            except (ValueError, KeyError):
                continue
            aw.setdefault(asin, {})[wi] = {'rev': rev, 'u': u, 'pr': pr, 'pu': pu}
            titles[asin] = row.get('title', '')

print(f'ASINs con ventas: {len(aw)} / {len(catalog)} ({len(catalog)-len(aw)} sin ventas en las 4 semanas)')

# ── 3. Agregación ──────────────────────────────────────────────────
def agg(asins):
    W = [{'rev': 0, 'u': 0, 'pr': 0, 'pu': 0, 'has': False} for _ in range(4)]
    for a in asins:
        for wi in range(4):
            d = (aw.get(a) or {}).get(wi)
            if d and (d['u'] > 0 or d['rev'] > 0):
                W[wi]['rev'] += d['rev']
                W[wi]['u']   += d['u']
                W[wi]['pr']  += d['pr']
                W[wi]['pu']  += d['pu']
                W[wi]['has']  = True
    total_u = sum(w['u'] for w in W)
    total_r = sum(w['rev'] for w in W)
    days    = sum(7 for w in W if w['has'])
    mu = total_u / days if days else 0
    mr = total_r / days if days else 0
    prev3 = [W[i] for i in range(3) if W[i]['has']]
    t = None
    if W[3]['has'] and prev3:
        avg = sum(w['u'] for w in prev3) / len(prev3)
        if avg > 0:
            t = (W[3]['u'] - avg) / avg * 100
    return {'W': W, 'mu': mu, 'mr': mr, 'tend': t}

# ── 4. Estructura familia → subfamilia → ASIN ──────────────────────
struct = {}
for asin, m in catalog.items():
    cat = m['col'] or m['familia']
    struct.setdefault(cat, {}).setdefault(m['sub'], [])
    if asin not in struct[cat][m['sub']]:
        struct[cat][m['sub']].append(asin)

EXPLICIT_CATS = set(CAT_ORDER) - {'_RESTO_'}
resto_subs = {}
for cat_key, subs in struct.items():
    if cat_key not in EXPLICIT_CATS:
        for sub, asins in subs.items():
            resto_subs.setdefault(f'{cat_key} — {sub}', []).extend(asins)
struct['_RESTO_'] = resto_subs

dash = {}
for cat in CAT_ORDER:
    if cat not in struct:
        continue
    display_name = 'Resto' if cat == '_RESTO_' else cat
    all_a = [a for subs in struct[cat].values() for a in subs]
    if not any(a in aw for a in all_a):
        continue
    ca = agg(all_a)
    subs_d = {}
    for sub, asins in sorted(struct[cat].items()):
        if not any(a in aw for a in asins):
            continue
        sa = agg(asins)
        ad = {}
        for a in sorted(asins, key=lambda x: -(aw.get(x, {}).get(3, {'u': 0})['u'])):
            if a not in aw:
                continue
            ad[a] = {
                'desc': (catalog[a]['desc'] or titles.get(a, ''))[:48],
                'agg':  agg([a]),
            }
        if ad:
            subs_d[sub] = {'agg': sa, 'asins': ad}
    if subs_d:
        dash[cat] = {'agg': ca, 'subs': subs_d, 'display': display_name}

# ── 5. Datos para gráficos ─────────────────────────────────────────
week_totals_u = [0, 0, 0, 0]
week_totals_r = [0.0, 0.0, 0.0, 0.0]
cat_chart_data = []

for cat in CAT_ORDER:
    if cat not in dash:
        continue
    d = dash[cat]['agg']
    wu = [d['W'][i]['u']   for i in range(4)]
    wr = [round(d['W'][i]['rev'], 0) for i in range(4)]
    for i in range(4):
        week_totals_u[i] += d['W'][i]['u']
        week_totals_r[i] += d['W'][i]['rev']
    n_active = sum(len(subs['asins']) for subs in dash[cat]['subs'].values())
    cat_chart_data.append({'name': dash[cat].get('display', cat), 'u': wu, 'r': wr, 'active': n_active})

total_u   = sum(week_totals_u)
total_r   = sum(week_totals_r)
best_wi   = week_totals_u.index(max(week_totals_u))
WEEK_NAMES = [w['label'] for w in WEEKS]

# ── 6. Helpers HTML ────────────────────────────────────────────────
def fu(v):
    if v == 0:  return '—'
    if v >= 1000: return f'{v/1000:.1f}k'
    return str(int(round(v)))

def fr(v):
    if v == 0:  return ''
    if v >= 1000: return f'{v/1000:.1f}k€'
    return f'{int(round(v))}€'

def fmed(u): return f'{int(round(u))}/d' if u > 0 else '—'

def tend_html(t):
    if t is None:   return '<span class="t-na">—</span>'
    if abs(t) < 3:  return f'<span class="t-flat">≈{t:+.0f}%</span>'
    if t > 0:       return f'<span class="t-up">↑{t:.0f}%</span>'
    return f'<span class="t-dn">↓{abs(t):.0f}%</span>'

def yoy_span(cur, prev, label, cls_prefix):
    if not prev or prev < 5 or not cur: return ''
    pct = (cur - prev) / prev * 100
    if abs(pct) > 499: return ''
    sign = '+' if pct >= 0 else ''
    cls  = 'yb-up' if pct >= 0 else 'yb-dn'
    return f'<span class="{cls_prefix} {cls}">{sign}{pct:.0f}%&nbsp;{label}</span>'

def week_cells(a):
    W = a['W']; mu = a['mu']; mr = a['mr']
    out = ''
    for w in W:
        cc = 'cn'
        if w['u'] > 0 and mu > 0:
            p  = w['u'] / mu
            cc = 'cg' if p >= 0.9 else ('cy' if p >= 0.6 else 'cr')
        yoy_u = yoy_span(w['u'],   w['pu'], 'uds', 'yy-u')
        yoy_r = yoy_span(w['rev'], w['pr'], '€',   'yy-r')
        yoy_block = f'<div class="wy">{yoy_u}{yoy_r}</div>' if (yoy_u or yoy_r) else ''
        out += f'''<td class="wk"><div class="wi {cc}">
          <div class="wu">{fu(w["u"])}</div>
          <div class="wr">{fr(w["rev"])}</div>
          {yoy_block}
        </div></td>'''
    out += f'<td class="tend">{tend_html(a["tend"])}</td>'
    return out

def build_rows():
    html = ''
    for cat in CAT_ORDER:
        if cat not in dash: continue
        d = dash[cat]; is_col = cat in COLS
        display  = d.get('display', cat)
        n_a = sum(len(v['asins']) for v in d['subs'].values())
        cat_id = cat.replace(' ', '_').replace('/', '_')
        html += f'''<tr class="rc {'col' if is_col else ''}" data-cat="{cat_id}">
          <td class="nc" onclick="tgC('{cat_id}')">
            <span class="ci">{'◆' if is_col else '▸'}</span>
            <span class="cn-lbl">{display}</span>
            <span class="meta">{len(d["subs"])} subfamilias · {n_a} ASINs</span>
          </td>
          <td class="mc"><div class="med-pill">{fmed(d["agg"]["mu"])}</div></td>
          {week_cells(d["agg"])}
        </tr>'''
        for sub, sd in d['subs'].items():
            sub_id = sub.replace("'", "").replace('"', '').replace(' ', '_')
            html += f'''<tr class="rs" data-p="{cat_id}" style="display:none">
              <td class="ns" onclick="tgS('{cat_id}','{sub_id}')">
                <span class="si">▹</span><span class="sn-lbl">{sub}</span>
                <span class="meta">{len(sd["asins"])} ASINs</span>
              </td>
              <td class="mc"><div class="med-pill sm">{fmed(sd["agg"]["mu"])}</div></td>
              {week_cells(sd["agg"])}
            </tr>'''
            for asin, ad in sd['asins'].items():
                html += f'''<tr class="ra" data-p="{cat_id}" data-s="{sub_id}" style="display:none">
                  <td class="an" title="{asin} — {ad["desc"]}"><span class="ac">{asin}</span>{ad["desc"]}</td>
                  <td class="mc"><span class="med-u">{fmed(ad["agg"]["mu"])}</span></td>
                  {week_cells(ad["agg"])}
                </tr>'''
    return html

rows_html = build_rows()
wk_ths = ''.join(
    f'<th class="wkh"><div class="wkl">{w["wk"]}</div><div class="wkd">{w["label"]}</div></th>'
    for w in WEEKS
)

# ── 7. JSON para gráficos ──────────────────────────────────────────
chart_js_data = json.dumps({
    'weekNames':    WEEK_NAMES,
    'weekTotalsU':  week_totals_u,
    'weekTotalsR':  [round(v) for v in week_totals_r],
    'cats':         cat_chart_data,
})

# ── 8. HTML ────────────────────────────────────────────────────────
html = f'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{CLIENT_NAME} · Semanas · {MARKETPLACE_LABEL}</title>
<style>
:root{{
  --bg:#090c14; --s1:#0f1320; --s2:#151929; --s3:#1d2235; --s4:#252b40;
  --bd:#2a2f4a; --bd2:#1a1f33;
  --t1:#dde3f5; --t2:#8892b0; --t3:#4a5270;
  --acc:#ff9900; --acc2:#ffb340;
  --g:#22c55e; --gb:rgba(34,197,94,.09);  --gbd:rgba(34,197,94,.4);
  --y:#f59e0b; --yb:rgba(245,158,11,.09); --ybd:rgba(245,158,11,.4);
  --r:#ef4444; --rb:rgba(239,68,68,.09);  --rbd:rgba(239,68,68,.4);
  --ff:'Segoe UI',system-ui,-apple-system,sans-serif;
  --ff-mono:'Cascadia Code','Consolas',monospace;
}}
@media(prefers-color-scheme:light){{:root{{
  --bg:#f0f2f8; --s1:#fff; --s2:#f7f9fe; --s3:#edf0f8; --s4:#e4e8f4;
  --bd:#cdd2e8; --bd2:#dde1f0; --t1:#131728; --t2:#556080; --t3:#8898b8;
}}}}
:root[data-theme="dark"]{{--bg:#090c14;--s1:#0f1320;--s2:#151929;--s3:#1d2235;--s4:#252b40;--bd:#2a2f4a;--bd2:#1a1f33;--t1:#dde3f5;--t2:#8892b0;--t3:#4a5270;}}
:root[data-theme="light"]{{--bg:#f0f2f8;--s1:#fff;--s2:#f7f9fe;--s3:#edf0f8;--s4:#e4e8f4;--bd:#cdd2e8;--bd2:#dde1f0;--t1:#131728;--t2:#556080;--t3:#8898b8;}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--t1);font-family:var(--ff);font-size:12px}}
.hdr{{background:var(--s2);border-bottom:1px solid var(--bd);padding:12px 20px;display:flex;align-items:center;gap:14px}}
.hdr-logo{{width:32px;height:32px;border-radius:7px;background:linear-gradient(135deg,#ff9900,#e07000);display:flex;align-items:center;justify-content:center;font-size:16px;font-weight:900;color:#fff;flex-shrink:0;box-shadow:0 2px 10px rgba(255,153,0,.25)}}
.hdr-t1{{font-size:14px;font-weight:700;color:var(--t1)}}
.hdr-t2{{font-size:10px;color:var(--t3);margin-top:1px}}
.badge-es{{background:rgba(255,153,0,.12);border:1px solid rgba(255,153,0,.25);color:var(--acc2);font-size:10px;font-weight:600;padding:2px 9px;border-radius:20px;margin-left:auto}}
.btn-sm{{background:var(--s3);border:1px solid var(--bd);color:var(--t2);padding:4px 9px;border-radius:5px;cursor:pointer;font-size:11px;transition:all .15s}}
.btn-sm:hover{{background:var(--s4);color:var(--t1)}}
.btn-sm.on{{background:rgba(255,153,0,.13);border-color:rgba(255,153,0,.35);color:var(--acc)}}
.summary{{background:var(--s1);border-bottom:1px solid var(--bd);padding:16px 20px;display:grid;gap:14px}}
.kpi-row{{display:flex;gap:10px;flex-wrap:wrap}}
.kpi{{background:var(--s2);border:1px solid var(--bd);border-radius:8px;padding:10px 14px;min-width:130px;flex:1}}
.kpi-val{{font-size:22px;font-weight:700;color:var(--t1);font-variant-numeric:tabular-nums;letter-spacing:-.5px;line-height:1}}
.kpi-lbl{{font-size:10px;color:var(--t3);text-transform:uppercase;letter-spacing:.6px;margin-top:4px}}
.kpi-sub{{font-size:10px;color:var(--t2);margin-top:2px}}
.kpi-acc .kpi-val{{color:var(--acc2)}}
.kpi-g .kpi-val{{color:var(--g)}}
.chart-area{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
.chart-box{{background:var(--s2);border:1px solid var(--bd);border-radius:8px;padding:12px;}}
.chart-title{{font-size:10px;font-weight:600;color:var(--t3);text-transform:uppercase;letter-spacing:.6px;margin-bottom:10px}}
canvas{{display:block;width:100%;}}
.toolbar{{background:var(--s1);border-bottom:1px solid var(--bd2);padding:7px 20px;display:flex;align-items:center;gap:8px;flex-wrap:wrap}}
.tlbl{{font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:.7px;font-weight:600}}
.tbar-sep{{width:1px;height:18px;background:var(--bd);margin:0 3px}}
.leg{{display:flex;align-items:center;gap:5px;font-size:10px;color:var(--t2)}}
.leg-b{{width:3px;height:12px;border-radius:1px;flex-shrink:0}}
.wrap{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;min-width:780px}}
thead th{{background:var(--s2);color:var(--t3);font-size:9px;font-weight:600;text-transform:uppercase;letter-spacing:.7px;padding:8px 10px;border-bottom:1px solid var(--bd);white-space:nowrap}}
th.nh{{text-align:left;min-width:260px}}
th.mh{{min-width:62px;text-align:center}}
th.wkh{{min-width:82px;text-align:center;border-left:1px solid var(--bd2)}}
.wkl{{font-size:11px;font-weight:700;color:var(--t2)}}
.wkd{{font-size:9px;color:var(--t3);margin-top:1px}}
th.tndh{{min-width:68px;text-align:center}}
tr.rc td{{background:var(--s2);border-bottom:1px solid var(--bd);padding:10px 12px;transition:background .1s;cursor:pointer}}
tr.rc td.nc{{border-left:4px solid transparent}}
tr.rc:hover td{{background:var(--s3)}}
tr.rc.col td{{background:linear-gradient(90deg,rgba(255,153,0,.06) 0%,var(--s2) 50%)!important}}
tr.rc.col td.nc{{border-left:4px solid var(--acc)!important}}
tr.rc.col:hover td{{background:linear-gradient(90deg,rgba(255,153,0,.11) 0%,var(--s3) 50%)!important}}
.ci{{font-size:9px;color:var(--t3);margin-right:7px;transition:transform .15s;display:inline-block}}
.rc.open .ci{{transform:rotate(90deg);color:var(--acc)}}
.cn-lbl{{font-size:13px;font-weight:700}}
.meta{{font-size:9px;color:var(--t3);margin-left:6px}}
.nc{{display:flex;align-items:center;user-select:none}}
tr.rs td{{background:var(--s1);border-bottom:1px solid var(--bd2);padding:7px 12px 7px 26px;transition:background .1s}}
tr.rs:hover td{{background:var(--s2)}}
tr.rs td:first-child{{border-left:2px solid var(--bd)}}
.si{{font-size:9px;color:var(--t3);margin-right:5px}}
.sn-lbl{{font-size:11px;font-weight:600;color:var(--t2)}}
.ns{{display:flex;align-items:center;cursor:pointer;user-select:none}}
tr.rs.open .si{{color:var(--acc2)}}
tr.ra td{{background:var(--bg);border-bottom:1px solid rgba(26,31,51,.6);padding:5px 12px 5px 38px;vertical-align:middle}}
tr.ra:hover td{{background:var(--s1)}}
tr.ra td:first-child{{border-left:2px solid var(--bd2)}}
.an{{font-size:10px;color:var(--t2);max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.ac{{font-family:var(--ff-mono);font-size:9px;color:var(--t3);margin-right:5px}}
.mc{{text-align:center;vertical-align:middle;padding:4px 6px!important}}
.med-pill{{display:inline-block;background:var(--s3);border:1px solid var(--bd);color:var(--t2);font-size:10px;font-weight:600;padding:2px 7px;border-radius:10px;white-space:nowrap}}
.med-pill.sm{{font-size:9px;padding:1px 6px;background:var(--s2)}}
.med-u{{font-size:9px;color:var(--t3);font-weight:600}}
td.wk{{padding:2px 2px;border-left:1px solid var(--bd2);vertical-align:middle;text-align:center}}
.wi{{border-radius:6px;padding:6px 5px;display:flex;flex-direction:column;align-items:center;gap:1px;border-left:3px solid transparent;transition:background .1s}}
tr.rc td.wk .wi{{padding:8px 6px;border-radius:8px}}
.wi.cg{{background:var(--gb);border-left-color:var(--gbd)}}
.wi.cy{{background:var(--yb);border-left-color:var(--ybd)}}
.wi.cr{{background:var(--rb);border-left-color:var(--rbd)}}
.wu{{font-size:15px;font-weight:700;line-height:1;font-variant-numeric:tabular-nums}}
tr.rc .wu{{font-size:18px}}
.wi.cg .wu{{color:var(--g)}} .wi.cy .wu{{color:var(--y)}} .wi.cr .wu{{color:var(--r)}} .wi.cn .wu{{color:var(--t3)}}
.wr{{font-size:9px;color:var(--t3);font-weight:400;margin-top:1px}}
tr.rc .wr{{font-size:10px}}
.wy{{margin-top:2px;display:flex;gap:2px;justify-content:center}}
.yy-u,.yy-r{{font-size:8px;font-weight:700;padding:1px 4px;border-radius:3px;white-space:nowrap}}
.yy-r{{display:none}}
.yb-up{{background:rgba(34,197,94,.13);color:#4ade80}}
.yb-dn{{background:rgba(239,68,68,.11);color:#f87171}}
td.tend{{text-align:center;padding:3px 7px;vertical-align:middle}}
.t-up{{font-size:12px;font-weight:700;color:var(--g)}}
.t-dn{{font-size:12px;font-weight:700;color:var(--r)}}
.t-flat{{font-size:11px;color:var(--t3)}}
.t-na{{font-size:11px;color:var(--t3)}}
.ft{{padding:10px 20px;border-top:1px solid var(--bd2);font-size:9px;color:var(--t3);display:flex;gap:12px;flex-wrap:wrap}}
</style>
</head>
<body>

<div class="hdr">
  <div class="hdr-logo">{CLIENT_LETTER}</div>
  <div>
    <div class="hdr-t1">{CLIENT_NAME} · Seguimiento Semanal Amazon Vendor</div>
    <div class="hdr-t2">Sell-out · ASINs activos · últimas 4 semanas · {MARKETPLACE_LABEL} · MerchantSpring</div>
  </div>
  <span class="badge-es">{MARKETPLACE_FLAG} {MARKETPLACE_LABEL}</span>
  <button class="btn-sm" id="btnTh" onclick="tgTh()">🌙</button>
</div>

<div class="summary">
  <div class="kpi-row">
    <div class="kpi kpi-acc">
      <div class="kpi-val">{round(total_u):,}</div>
      <div class="kpi-lbl">Uds vendidas · 4 semanas</div>
      <div class="kpi-sub">Mejor: {WEEK_NAMES[best_wi]} ({round(week_totals_u[best_wi]):,} uds)</div>
    </div>
    <div class="kpi">
      <div class="kpi-val">{round(total_r/1000,1)}k€</div>
      <div class="kpi-lbl">Facturación · 4 semanas</div>
      <div class="kpi-sub">Media/día: {round(total_r/28):,}€</div>
    </div>
    <div class="kpi kpi-g">
      <div class="kpi-val">{len(aw):,}</div>
      <div class="kpi-lbl">ASINs activos</div>
      <div class="kpi-sub">{len(catalog)-len(aw):,} sin ventas en el periodo</div>
    </div>
    <div class="kpi">
      <div class="kpi-val">{len(catalog):,}</div>
      <div class="kpi-lbl">ASINs catálogo</div>
      <div class="kpi-sub">Anualidades activas incluidas</div>
    </div>
  </div>
  <div class="chart-area">
    <div class="chart-box">
      <div class="chart-title">Evolución semanal · unidades totales</div>
      <canvas id="cvWeek" height="110"></canvas>
    </div>
    <div class="chart-box">
      <div class="chart-title">Tendencia por categoría · unidades (normalizado)</div>
      <canvas id="cvCat" height="110"></canvas>
    </div>
  </div>
</div>

<div class="toolbar">
  <span class="tlbl">Semáforo</span>
  <div class="leg"><div class="leg-b" style="background:var(--g)"></div>≥90% media</div>
  <div class="leg"><div class="leg-b" style="background:var(--y)"></div>60–89%</div>
  <div class="leg"><div class="leg-b" style="background:var(--r)"></div>&lt;60%</div>
  <div class="tbar-sep"></div>
  <span class="tlbl">Ver</span>
  <button class="btn-sm on" id="bRev" onclick="togRev()">€ Facturación</button>
  <button class="btn-sm on" id="bU"   onclick="togU()">📦 Unidades</button>
  <div class="tbar-sep"></div>
  <span class="tlbl">YOY</span>
  <div class="leg" style="gap:3px"><span style="background:rgba(34,197,94,.13);color:#4ade80;font-size:8px;font-weight:700;padding:1px 4px;border-radius:3px">+X%</span>vs misma semana año anterior</div>
  <div class="tbar-sep"></div>
  <button class="btn-sm" onclick="expAll()">+ Todo</button>
  <button class="btn-sm" onclick="colAll()">− Todo</button>
  <span style="margin-left:auto;font-size:9px;color:var(--t3)">Actualizado {UPDATE_DATE} · Tendencia: W4 vs media W1–W3 en unidades</span>
</div>

<div class="wrap">
<table>
  <thead>
    <tr>
      <th class="nh">Categoría / Subfamilia / ASIN</th>
      <th class="mh">Med/día</th>
      {wk_ths}
      <th class="tndh">Tend.</th>
    </tr>
  </thead>
  <tbody id="tb">{rows_html}</tbody>
</table>
</div>
<div class="ft">
  <span>Fuente: MerchantSpring · Vendor manufacturing view</span>
  <span>·</span><span>Semáforo vs media/día del periodo</span>
  <span>·</span><span>YOY = misma semana año anterior (oculto si prev &lt;5 uds)</span>
</div>

<script>
const CD = {chart_js_data};
let sRev=true, sU=true;

function tgTh(){{
  const r=document.documentElement,c=r.getAttribute('data-theme')||'dark';
  r.setAttribute('data-theme',c==='dark'?'light':'dark');
  document.getElementById('btnTh').textContent=c==='dark'?'☀':'🌙';
  setTimeout(drawCharts,30);
}}

function togRev(){{
  sRev=!sRev;
  document.getElementById('bRev').classList.toggle('on',sRev);
  document.querySelectorAll('.wr').forEach(e=>e.style.display=sRev?'':'none');
  updateYoy();
}}
function togU(){{
  sU=!sU;
  document.getElementById('bU').classList.toggle('on',sU);
  document.querySelectorAll('.wu').forEach(e=>e.style.display=sU?'':'none');
  updateYoy();
}}
function updateYoy(){{
  document.querySelectorAll('.yy-u').forEach(e=>e.style.display=sU?'':'none');
  document.querySelectorAll('.yy-r').forEach(e=>e.style.display=(!sU&&sRev)?'':'none');
}}

function tgC(cat){{
  const rows=[...document.querySelectorAll('[data-p="'+cat+'"]')];
  const op=rows[0]&&rows[0].style.display==='none';
  rows.forEach(tr=>tr.style.display=op?'':'none');
  const r=[...document.querySelectorAll('tr.rc')].find(r=>r.dataset.cat===cat);
  if(r)r.classList.toggle('open',op);
}}
function tgS(cat,sub){{
  const rows=[...document.querySelectorAll('[data-p="'+cat+'"][data-s="'+sub+'"]')];
  const op=rows[0]&&rows[0].style.display==='none';
  rows.forEach(tr=>tr.style.display=op?'':'none');
  const r=[...document.querySelectorAll('tr.rs')].find(r=>r.dataset.p===cat&&r.dataset.s===sub);
  if(r)r.classList.toggle('open',op);
}}
function expAll(){{document.querySelectorAll('.rs,.ra').forEach(r=>r.style.display='');}}
function colAll(){{document.querySelectorAll('.rs,.ra').forEach(r=>r.style.display='none');}}

function getCSS(v){{return getComputedStyle(document.documentElement).getPropertyValue(v).trim();}}

function drawWeekChart(canvas){{
  const dpr=window.devicePixelRatio||1;
  const W=canvas.offsetWidth, H=canvas.offsetHeight||110;
  canvas.width=W*dpr; canvas.height=H*dpr;
  const ctx=canvas.getContext('2d'); ctx.scale(dpr,dpr);
  const vals=CD.weekTotalsU, n=vals.length;
  const maxV=Math.max(...vals)*1.15;
  const pad={{l:36,r:10,t:8,b:28}};
  const bw=Math.floor((W-pad.l-pad.r)/n*0.55);
  const gap=(W-pad.l-pad.r-bw*n)/(n+1);
  const t1=getCSS('--t1'), t2=getCSS('--t2'), t3=getCSS('--t3'), bd=getCSS('--bd');
  const acc=getCSS('--acc');
  ctx.strokeStyle=bd; ctx.lineWidth=.5;
  for(let i=0;i<=4;i++){{
    const y=pad.t+(H-pad.t-pad.b)/4*i;
    ctx.beginPath(); ctx.moveTo(pad.l,y); ctx.lineTo(W-pad.r,y); ctx.stroke();
    if(i<4){{ctx.fillStyle=t3;ctx.font='9px sans-serif';ctx.textAlign='right';
      ctx.fillText(Math.round(maxV/4*(4-i)),pad.l-4,y+3);}}
  }}
  vals.forEach((v,i)=>{{
    const x=pad.l+gap*(i+1)+bw*i;
    const bh=(v/maxV)*(H-pad.t-pad.b);
    const y=H-pad.b-bh;
    const isBest=(i===CD.weekTotalsU.indexOf(Math.max(...CD.weekTotalsU)));
    ctx.fillStyle=isBest?acc:'rgba(255,153,0,.45)';
    ctx.beginPath(); ctx.roundRect(x,y,bw,bh,3); ctx.fill();
    ctx.fillStyle=isBest?acc:t2; ctx.font='bold 9px sans-serif'; ctx.textAlign='center';
    ctx.fillText(v>=1000?(v/1000).toFixed(1)+'k':v, x+bw/2, y-3);
    ctx.fillStyle=t3; ctx.font='9px sans-serif';
    ctx.fillText(CD.weekNames[i].split('–')[0].trim(),x+bw/2,H-pad.b+11);
  }});
}}

function drawCatChart(canvas){{
  const dpr=window.devicePixelRatio||1;
  const W=canvas.offsetWidth||300, H=canvas.offsetHeight||110;
  canvas.width=W*dpr; canvas.height=H*dpr;
  const ctx=canvas.getContext('2d'); ctx.scale(dpr,dpr);
  const isDark=(document.documentElement.getAttribute('data-theme')||'dark')==='dark';
  const cats=CD.cats.filter(c=>c.u.some(v=>v>0));
  const n=cats.length, WEEKS=4;
  const NAME_W=76, PAD_T=16, PAD_B=4, PAD_R=6;
  const CELL_W=Math.floor((W-NAME_W-PAD_R)/WEEKS);
  const CELL_H=Math.floor((H-PAD_T-PAD_B)/n);
  const t2=isDark?'#8892b0':'#556080';
  const t3=isDark?'#4a5270':'#8898b8';
  const bd=isDark?'rgba(255,255,255,.06)':'rgba(0,0,0,.06)';
  ['W1','W2','W3','W4'].forEach((lbl,wi)=>{{
    const cx=NAME_W+wi*CELL_W+CELL_W/2;
    ctx.fillStyle=t3; ctx.font='bold 9px sans-serif'; ctx.textAlign='center';
    ctx.fillText(lbl,cx,PAD_T-4);
  }});
  cats.forEach((cat,ci)=>{{
    const maxU=Math.max(...cat.u)||1;
    const ry=PAD_T+ci*CELL_H;
    ctx.fillStyle=t2; ctx.font='10px sans-serif'; ctx.textAlign='right';
    const lbl=cat.name.length>10?cat.name.slice(0,9)+'…':cat.name;
    ctx.fillText(lbl, NAME_W-4, ry+CELL_H/2+3.5);
    for(let wi=0;wi<WEEKS;wi++){{
      const ratio=cat.u[wi]/maxU;
      const alpha=0.12+ratio*0.78;
      ctx.fillStyle=`rgba(255,153,0,${{alpha.toFixed(2)}})`;
      const cx=NAME_W+wi*CELL_W, cy=ry;
      ctx.fillRect(cx+1,cy+1,CELL_W-2,CELL_H-2);
    }}
    if(ci<n-1){{
      ctx.strokeStyle=bd; ctx.lineWidth=.5;
      ctx.beginPath(); ctx.moveTo(0,ry+CELL_H); ctx.lineTo(W,ry+CELL_H); ctx.stroke();
    }}
  }});
}}

function drawCharts(){{
  const cw=document.getElementById('cvWeek'), cc=document.getElementById('cvCat');
  if(cw.offsetWidth>0){{drawWeekChart(cw); drawCatChart(cc);}}
}}
CD.cats.forEach(c=>{{ c.mu_all = c.u.reduce((a,b)=>a+b,0)/4 || 1; }});
window.addEventListener('load',drawCharts);
window.addEventListener('resize',drawCharts);
</script>
</body>
</html>'''

with open(OUT_PATH, 'w', encoding='utf-8') as f:
    f.write(html)
print(f'\n✓ Dashboard: {OUT_PATH} ({len(html):,} bytes)')
print(f'\n=== RESUMEN ===')
print(f'ASINs catálogo:        {len(catalog):,}')
print(f'ASINs con ventas:      {len(aw):,}')
print(f'ASINs sin ventas:      {len(catalog)-len(aw):,}')
print(f'Total uds 4 semanas:   {total_u:,}')
print(f'Total facturación:     {round(total_r/1000,1)}k€')
