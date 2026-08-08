"""
auto_update.py — Actualización automática semanal FINOCAM_Weekly.html
Se ejecuta cada lunes desde el agente cloud (o manualmente).

Flujo:
  1. Calcula las 4 semanas ISO completas anteriores al lunes de hoy
  2. Descarga los 4 CSVs desde MerchantSpring via MCP
  3. Genera FINOCAM_Weekly.html con build_weekly.py
  4. Hace git push a GitHub Pages
  5. Envía email resumen via SendGrid

Variables de entorno necesarias:
  SENDGRID_API_KEY  — clave API de SendGrid
  GITHUB_PAT        — Personal Access Token con permiso repo
  GITHUB_REPO       — dparodi-roicos/Finocam-Amazon
  EMAIL_FROM        — dparodi@roicos.com
  EMAIL_TO          — dparodi@roicos.com
"""

import os, sys, subprocess, datetime, json, csv, io, urllib.request, urllib.error
sys.stdout.reconfigure(encoding='utf-8')

# ── CONFIG ──────────────────────────────────────────────────────────
CHANNEL_ID  = '21907968'
MERCHANT_ID = 'amzn1.vg.3250111'
REPO        = os.environ.get('GITHUB_REPO', 'dparodi-roicos/Finocam-Amazon')
EMAIL_FROM  = os.environ.get('EMAIL_FROM',  'dparodi@roicos.com')
EMAIL_TO    = os.environ.get('EMAIL_TO',    'dparodi@roicos.com')
SENDGRID_KEY = os.environ.get('SENDGRID_API_KEY', '')
GITHUB_PAT   = os.environ.get('GITHUB_PAT', '')

TOOLS_DIR  = os.path.dirname(os.path.abspath(__file__))
REPO_DIR   = os.path.dirname(TOOLS_DIR)
DATA_DIR   = os.path.join(TOOLS_DIR, 'data')
CATALOG_JSON = os.path.join(DATA_DIR, 'catalog_es.json')
OUT_HTML   = os.path.join(REPO_DIR, 'FINOCAM_Weekly.html')

os.makedirs(DATA_DIR, exist_ok=True)

# ── 1. Calcular las 4 semanas (lun-dom) anteriores al lunes de hoy ──
def get_4_weeks():
    today = datetime.date.today()
    # El lunes de esta semana (día 0 = lunes en isoweekday)
    this_monday = today - datetime.timedelta(days=today.weekday())
    weeks = []
    for i in range(4, 0, -1):
        mon = this_monday - datetime.timedelta(weeks=i)
        sun = mon + datetime.timedelta(days=6)
        weeks.append({
            'wk':    f'W{5-i}',
            'label': f'{mon.day}–{sun.day} {mon.strftime("%b")}',
            'start': mon,
            'end':   sun,
        })
    return weeks

WEEKS = get_4_weeks()
print('Semanas a procesar:')
for w in WEEKS:
    print(f'  {w["wk"]}: {w["label"]} ({w["start"]} → {w["end"]})')

# ── 2. Descargar CSVs desde MerchantSpring ──────────────────────────
# NOTA: esta sección usa el cliente REST de MerchantSpring directamente.
# Si se ejecuta desde Claude Code (con MCP disponible), se puede usar
# el MCP interactivamente. En modo automatizado necesita la API key.
# Por ahora: lee los CSVs si ya existen en DATA_DIR, si no avisa.
csv_paths = []
for wi, w in enumerate(WEEKS):
    path = os.path.join(DATA_DIR, f'weekly_{wi}.csv')
    if os.path.exists(path):
        csv_paths.append(path)
        print(f'  CSV {wi}: {path} (existente)')
    else:
        print(f'  AVISO: falta {path} — descarga manualmente o usa el MCP')
        csv_paths.append(None)

if any(p is None for p in csv_paths):
    print('\nATENCIÓN: faltan CSVs. El agente debe descargarlos via MerchantSpring MCP antes.')
    print('Ejecuta primero: fetch_merchantspring_data.py')
    sys.exit(1)

# ── 3. Generar HTML (reutiliza lógica de build_weekly.py) ────────────
# Importar y ejecutar el módulo de build
build_script = os.path.join(TOOLS_DIR, 'build_weekly.py')
# Actualizar CSV_DIR y OUT_PATH en el entorno antes de importar
os.environ['WEEKLY_CSV_DIR']  = DATA_DIR
os.environ['WEEKLY_OUT_PATH'] = OUT_HTML
result = subprocess.run([sys.executable, build_script], capture_output=True, text=True, encoding='utf-8')
print(result.stdout)
if result.returncode != 0:
    print('ERROR en build_weekly.py:', result.stderr)
    sys.exit(1)

# ── 4. Git push a GitHub Pages ───────────────────────────────────────
if GITHUB_PAT:
    remote_url = f'https://x-access-token:{GITHUB_PAT}@github.com/{REPO}.git'
    today_str  = datetime.date.today().isoformat()
    cmds = [
        ['git', '-C', REPO_DIR, 'add', 'FINOCAM_Weekly.html'],
        ['git', '-C', REPO_DIR, 'commit', '-m', f'auto: actualización semanal {today_str}'],
        ['git', '-C', REPO_DIR, 'push', remote_url, 'main'],
    ]
    for cmd in cmds:
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode != 0 and 'nothing to commit' not in r.stdout + r.stderr:
            print(f'  git error: {r.stderr[:200]}')
        else:
            print(f'  git ok: {" ".join(cmd[3:])}')
else:
    print('  Sin GITHUB_PAT — omitiendo push')

# ── 5. Email via SendGrid ────────────────────────────────────────────
def send_email(subject, body_html):
    if not SENDGRID_KEY:
        print('  Sin SENDGRID_API_KEY — omitiendo email')
        return
    payload = json.dumps({
        'personalizations': [{'to': [{'email': EMAIL_TO}]}],
        'from': {'email': EMAIL_FROM, 'name': 'Finocam · Roicos'},
        'subject': subject,
        'content': [{'type': 'text/html', 'value': body_html}],
    }).encode('utf-8')
    req = urllib.request.Request(
        'https://api.sendgrid.com/v3/mail/send',
        data=payload,
        headers={
            'Authorization': f'Bearer {SENDGRID_KEY}',
            'Content-Type':  'application/json',
        },
        method='POST',
    )
    try:
        with urllib.request.urlopen(req) as resp:
            print(f'  Email enviado → {EMAIL_TO} (status {resp.status})')
    except urllib.error.HTTPError as e:
        print(f'  Error SendGrid {e.code}: {e.read().decode()}')

week_range  = f'{WEEKS[0]["label"]} – {WEEKS[-1]["label"]}'
dashboard_url = f'https://dparodi-roicos.github.io/Finocam-Amazon/FINOCAM_Weekly.html'

send_email(
    subject=f'Finocam · Dashboard semanal actualizado — {week_range}',
    body_html=f'''
    <p>Hola,</p>
    <p>El dashboard semanal de <strong>Finocam Amazon ES</strong> ha sido actualizado
    con los datos de las semanas <strong>{week_range}</strong>.</p>
    <p><a href="{dashboard_url}" style="background:#ff9900;color:#fff;padding:10px 20px;
    border-radius:5px;text-decoration:none;font-weight:bold;">Ver dashboard →</a></p>
    <p style="color:#888;font-size:12px;margin-top:20px">
    Generado automáticamente · {datetime.date.today().isoformat()} ·
    Fuente: MerchantSpring Vendor ES</p>
    ''',
)

print('\n✓ Actualización completada.')
