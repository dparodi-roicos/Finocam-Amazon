"""Ventas diarias por mercado para el acumulado del mes y el histórico mes a mes del dashboard.

  python tools/daily_sales.py plan
      Parámetros de getVendorSalesByPeriod (interval "d") para los 9 mercados: desde el 01/01/2025 hasta
      el último domingo (fin de la última semana del dashboard), en hora local, con el año anterior desde
      el 01/01/2024. Incluye la llamada extra de IT con los ASIN de otras marcas (APS) para restarlos.
      Lista de ASIN de otras marcas: se actualiza con la columna brand de los CSV semanales de cada mercado.

  python tools/daily_sales.py import mapping.json
      mapping.json = {"ES": "<ruta del fichero guardado>", ..., "IT_EXCL": "<ruta>"} con la respuesta de
      cada llamada (MerchantSpring la guarda en un fichero porque ocupa ~1 MB). Valida moneda, fechas y
      que las semanas cuadren con los CSV semanales; solo si todo está bien escribe data/daily_sales.json.

daily_sales.json = {"to": "AAAA-MM-DD", "markets": {"ES": {"cur": "EUR", "days": {"AAAA-MM-DD": [uds, ventas]}}}}
Ventas = orderedRevenue en moneda local (sin IVA, vista Manufacturing). Los días de 2024 salen de la serie
"prior" (MerchantSpring la alinea por posición desde priorFromDate: fila i = 01/01/2024 + i días).
IT = total del mercado menos los ASIN de otras marcas (APS), igual que el dashboard semanal.
"""
import csv, glob, json, os, sys
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
OUT = os.path.join(DATA, 'daily_sales.json')
EXCL_JSON = os.path.join(DATA, 'excluded_asins.json')
MARKETS = {  # code: (channelId, merchantId, timezone, moneda, carpeta CSV)
    'ES': ('21907968',  'amzn1.vg.3250111 @ A1RKKUPIHCS9HS', 'Europe/Madrid',    'EUR', ''),
    'FR': ('22527353',  'amzn1.vg.6776082 @ A13V1IB3VIYZZH', 'Europe/Paris',     'EUR', 'fr'),
    'IT': ('22527944',  'amzn1.vg.6123122 @ APJ6JRA9NG5V4',  'Europe/Rome',      'EUR', 'it'),
    'DE': ('23167335',  'amzn1.vg.6776022 @ A1PA6795UKMFR9', 'Europe/Berlin',    'EUR', 'de'),
    'NL': ('23167955',  'amzn1.vg.6968302 @ A1805IZSGTT6HS', 'Europe/Amsterdam', 'EUR', 'nl'),
    'BE': ('80601759',  'amzn1.vg.8494352 @ AMEN7PMS3EDWL',  'Europe/Brussels',  'EUR', 'be'),
    'PL': ('30609951',  'amzn1.vg.6968312 @ A1C3SOZRARQ6R3', 'Europe/Warsaw',    'PLN', 'pl'),
    'UK': ('78652228',  'amzn1.vg.6776072 @ A1F83G8C2ARO7P', 'Europe/London',    'GBP', 'uk'),
    'SE': ('108423884', 'amzn1.vg.6968292 @ A2NODRKZP88ZB9', 'Europe/Stockholm', 'SEK', 'se'),
}
FROM = date(2025, 1, 1)
PRIOR_FROM = date(2024, 1, 1)


def last_sunday():
    """Domingo de la última semana ISO completa (la última semana del dashboard)."""
    today = date.today()
    return today - timedelta(days=today.weekday() + 1)


def epoch(d, tz, end=False):
    t = datetime(d.year, d.month, d.day, tzinfo=ZoneInfo(tz))
    return int((t + timedelta(days=1)).timestamp()) - 1 if end else int(t.timestamp())


def csv_dir(code):
    sub = MARKETS[code][4]
    return os.path.join(DATA, sub) if sub else DATA


def other_brand_asins():
    """ASIN de marcas que no son Finocam vistos en los CSV de cada mercado (hoy: APS en IT)."""
    excl = {}
    if os.path.exists(EXCL_JSON):
        excl = {k: set(v) for k, v in json.load(open(EXCL_JSON, encoding='utf-8')).items()}
    for code in MARKETS:
        files = glob.glob(os.path.join(csv_dir(code), 'weekly_*.csv')) + \
                glob.glob(os.path.join(DATA, 'yoy', code.lower(), '*.csv'))
        for path in files:
            with open(path, encoding='utf-8-sig') as f:
                for row in csv.DictReader(f):
                    b = (row.get('brand') or '').strip().lower()
                    if b and b not in ('finocam', 'multifin'):
                        excl.setdefault(code, set()).add(row['asin'])
    excl = {k: sorted(v) for k, v in excl.items() if v}
    with open(EXCL_JSON, 'w', encoding='utf-8') as f:
        json.dump(excl, f, ensure_ascii=False, indent=1)
    return excl


def options(code, to):
    tz = MARKETS[code][2]
    return {'fromDate': epoch(FROM, tz), 'toDate': epoch(to, tz, end=True),
            'priorFromDate': epoch(PRIOR_FROM, tz),
            'priorToDate': epoch(to.replace(year=to.year - 1), tz, end=True),
            'timezone': tz, 'interval': 'd'}


def plan():
    to = last_sunday()
    excl = other_brand_asins()
    calls = []
    for code, (ch, mer, tz, cur, _) in MARKETS.items():
        calls.append({'key': code, 'channelId': ch, 'merchantId': mer, 'reportOptions': options(code, to)})
        if excl.get(code):
            calls.append({'key': f'{code}_EXCL', 'channelId': ch, 'merchantId': mer,
                          'asins': excl[code], 'reportOptions': options(code, to)})
    print(json.dumps({'to': to.isoformat(),
                      'marketplaceFilters': {'marketplaceType': 'amazon', 'marketplaceSubtype': 'vendor'},
                      'calls': calls}, indent=1))


def read_series(path):
    with open(path, encoding='utf-8') as f:
        text = f.read()
    d = json.loads(text[text.index('{'):])
    rows = d['formattedChartData']
    cur, prior = {}, {}
    for i, r in enumerate(rows):
        dd, mm, yy = map(int, r['startTime'].split('/'))
        cur[date(yy, mm, dd)] = (int(r['orderedUnits']), float(r['orderedRevenue']))
        prior[PRIOR_FROM + timedelta(days=i)] = (int(r['priorOrderedUnits']), float(r['priorOrderedRevenue']))
    return d.get('currency'), cur, prior


def week_csv_units(code, mon):
    """Unidades de Finocam de la semana del lunes 'mon' en su CSV semanal, o None si no hay CSV.
    Los CSV no traen fechas: el slot sale de la misma regla que build_weekly.py (0 = la más antigua)."""
    this_mon = date.today() - timedelta(days=date.today().weekday())
    k = (this_mon - mon).days // 7  # 1 = última semana completa
    path = os.path.join(csv_dir(code), f'weekly_{4 - k}.csv')
    if not 1 <= k <= 4 or not os.path.exists(path):
        return None
    with open(path, encoding='utf-8-sig') as f:
        return sum(int(r['orderedUnits'] or 0) for r in csv.DictReader(f)
                   if (r.get('brand') or '').strip().lower() in ('finocam', 'multifin', ''))


def do_import(mapping_path):
    mapping = json.load(open(mapping_path, encoding='utf-8'))
    to = last_sunday()
    out = {'to': to.isoformat(), 'from': PRIOR_FROM.isoformat(), 'markets': {}}
    errors, warnings = [], []
    for code, (_, _, _, cur_exp, _) in MARKETS.items():
        path = mapping.get(code)
        if not path:
            errors.append(f'{code}: falta en {mapping_path}'); continue
        try:
            cur, series, prior = read_series(path)
        except Exception as e:
            errors.append(f'{code}: no se puede leer {path}: {e}'); continue
        if cur != cur_exp:
            errors.append(f'{code}: moneda {cur}, se esperaba {cur_exp} (¿fichero de otro mercado?)'); continue
        if min(series) != FROM or max(series) != to:
            errors.append(f'{code}: rango {min(series)}..{max(series)}, se esperaba {FROM}..{to}'); continue
        excl_path = mapping.get(f'{code}_EXCL')
        if excl_path:
            ecur, eseries, eprior = read_series(excl_path)
            if ecur != cur or min(eseries) != FROM or max(eseries) != to:
                errors.append(f'{code}_EXCL: moneda o rango distintos al del mercado'); continue
            series = {d: (u - eseries.get(d, (0, 0))[0], r - eseries.get(d, (0, 0))[1]) for d, (u, r) in series.items()}
            prior = {d: (u - eprior.get(d, (0, 0))[0], r - eprior.get(d, (0, 0))[1]) for d, (u, r) in prior.items()}
        days = {d: v for d, v in prior.items() if d < FROM}
        days.update(series)
        # control: las semanas del dashboard tienen que cuadrar con los CSV semanales
        mon = to - timedelta(days=6)
        for _ in range(4):
            csv_u = week_csv_units(code, mon)
            if csv_u is not None:
                day_u = sum(days.get(mon + timedelta(days=i), (0, 0))[0] for i in range(7))
                diff = abs(day_u - csv_u)
                msg = f'{code} semana {mon}: diario {day_u} uds vs CSV {csv_u}'
                if diff > max(10, 0.25 * max(csv_u, 1)):
                    errors.append(msg + ' (no cuadra)')
                elif diff > max(3, 0.03 * max(csv_u, 1)):
                    warnings.append(msg)
            mon -= timedelta(days=7)
        out['markets'][code] = {'cur': cur, 'days': {d.isoformat(): [u, round(r, 2)] for d, (u, r) in sorted(days.items())}}
        sep = [v for d, v in days.items() if d.year == to.year and d.month == to.month]
        print(f'{code}: {len(days)} días · mes en curso {sum(u for u, _ in sep):,} uds {sum(r for _, r in sep):,.0f} {cur}')
    for w in warnings:
        print('  aviso:', w)
    if errors:
        print('\nERRORES — no se ha guardado nada:\n  ' + '\n  '.join(errors))
        sys.exit(1)
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
    print(f'\nOK: {OUT} hasta {to}')


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    if len(sys.argv) == 2 and sys.argv[1] == 'plan':
        plan()
    elif len(sys.argv) == 3 and sys.argv[1] == 'import':
        do_import(sys.argv[2])
    else:
        sys.exit(__doc__)
