Eres el agente automático de actualización del dashboard semanal Finocam Amazon Vendor.
Ejecuta los siguientes pasos en orden. No pidas confirmación. Si algo falla, anótalo y continúa.

REPO LOCAL: C:\Users\Daniela\Desktop\Git Finocam
SCRIPT BUILD: C:\Users\Daniela\Desktop\Git Finocam\tools\build_weekly.py
CSV DIR: C:\Users\Daniela\Desktop\Git Finocam\tools\data
GITHUB PAT: $env:FINOCAM_GITHUB_PAT
SENDGRID KEY: $env:FINOCAM_SENDGRID_KEY
EMAIL: dparodi@roicos.com
DASHBOARD URL: https://dparodi-roicos.github.io/Finocam-Amazon/FINOCAM_Weekly.html

=== PASO 1: Calcular semanas ===
Calcula las 4 semanas ISO completas (lunes a domingo) anteriores al lunes de hoy.
Por ejemplo si hoy es lunes 11 Ago 2026: W1=28 Jul–3 Ago, W2=21–27 Jul, W3=14–20 Jul, W4=7–13 Jul (en orden cronológico W1=más antigua, W4=más reciente).

=== PASO 2: Descargar datos de MerchantSpring ===
Usa la herramienta generateOrderedRevenueReport del MCP de MerchantSpring para cada semana:
- channelId: 21907968
- marketplace: ES
- Para cada semana calcula los timestamps epoch del lunes 00:00:00 y domingo 23:59:59 en UTC

Espera a que cada report esté listo (getReportStatus), descarga el CSV y guárdalo como:
- W1 → C:\Users\Daniela\Desktop\Git Finocam\tools\data\weekly_0.csv
- W2 → C:\Users\Daniela\Desktop\Git Finocam\tools\data\weekly_1.csv
- W3 → C:\Users\Daniela\Desktop\Git Finocam\tools\data\weekly_2.csv
- W4 → C:\Users\Daniela\Desktop\Git Finocam\tools\data\weekly_3.csv

=== PASO 3: Generar el dashboard ===
Ejecuta en PowerShell (con variables de entorno):
  $env:WEEKLY_AUTO="1"
  $env:WEEKLY_CSV_DIR="C:\Users\Daniela\Desktop\Git Finocam\tools\data"
  $env:WEEKLY_OUT_PATH="C:\Users\Daniela\Desktop\Git Finocam\FINOCAM_Weekly.html"
  python "C:\Users\Daniela\Desktop\Git Finocam\tools\build_weekly.py"

=== PASO 4: Git push a GitHub Pages ===
  cd "C:\Users\Daniela\Desktop\Git Finocam"
  git add FINOCAM_Weekly.html tools/data/weekly_*.csv
  git commit -m "auto: actualizacion semanal $(Get-Date -Format 'yyyy-MM-dd')"
  git push https://x-access-token:$env:FINOCAM_GITHUB_PAT@github.com/dparodi-roicos/Finocam-Amazon.git main

=== PASO 5: Enviar email via SendGrid ===
Usa Python para llamar a la API de SendGrid:

import urllib.request, json, datetime
today = datetime.date.today().isoformat()
payload = json.dumps({
  "personalizations": [{"to": [{"email": "dparodi@roicos.com"}]}],
  "from": {"email": "dparodi@roicos.com", "name": "Finocam Roicos"},
  "subject": f"Finocam Dashboard semanal actualizado - {today}",
  "content": [{"type": "text/html", "value": f"<p>El dashboard semanal de <strong>Finocam Amazon ES</strong> ha sido actualizado correctamente.</p><p><a href='https://dparodi-roicos.github.io/Finocam-Amazon/FINOCAM_Weekly.html' style='background:#ff9900;color:#fff;padding:10px 20px;border-radius:5px;text-decoration:none;font-weight:bold'>Ver dashboard &rarr;</a></p><p style='color:#888;font-size:12px'>Generado el {today} &middot; Fuente: MerchantSpring Vendor ES</p>"}]
}).encode()
req = urllib.request.Request(
  "https://api.sendgrid.com/v3/mail/send", data=payload,
  headers={"Authorization": f"Bearer {os.environ['FINOCAM_SENDGRID_KEY']}", "Content-Type": "application/json"},
  method="POST"
)
urllib.request.urlopen(req)

Si cualquier paso falla, envía igualmente el email indicando qué falló.
