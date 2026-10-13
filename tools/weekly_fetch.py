"""Ayuda para la actualización semanal del dashboard (rutina de los lunes).

  python tools/weekly_fetch.py plan
      Las 4 últimas semanas ISO completas, con los epochs (lunes 00:00 - domingo 23:59:59 en hora
      local de cada mercado) y los IDs de MerchantSpring para generateOrderedRevenueReport, más las
      semanas del año anterior (YOY) que faltan en data/yoy/. Salida JSON.

  python tools/weekly_fetch.py download urls.json
      urls.json = {"es": {"0": url, "1": url, "2": url, "3": url, "yoy": {"2025_W42": url}}, "fr": {...}, ...}
      (url = downloadUrl de getReportStatus; null si el informe salió "errored" = sin datos).
      Descarga a una carpeta temporal y valida cabecera y rango lunes-domingo del nombre del CSV.
      Solo si todo está bien copia a data/<mercado>/weekly_N.csv y data/yoy/<mercado>/AAAA_Wss.csv.
"""
import csv, json, os, re, shutil, sys, tempfile, urllib.request
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
MARKETS = {  # code: (channelId, merchantId, timezone)
    'es': ('21907968',  'amzn1.vg.3250111 @ A1RKKUPIHCS9HS', 'Europe/Madrid'),
    'fr': ('22527353',  'amzn1.vg.6776082 @ A13V1IB3VIYZZH', 'Europe/Paris'),
    'it': ('22527944',  'amzn1.vg.6123122 @ APJ6JRA9NG5V4',  'Europe/Rome'),
    'de': ('23167335',  'amzn1.vg.6776022 @ A1PA6795UKMFR9', 'Europe/Berlin'),
    'nl': ('23167955',  'amzn1.vg.6968302 @ A1805IZSGTT6HS', 'Europe/Amsterdam'),
    'be': ('80601759',  'amzn1.vg.8494352 @ AMEN7PMS3EDWL',  'Europe/Brussels'),
    'pl': ('30609951',  'amzn1.vg.6968312 @ A1C3SOZRARQ6R3', 'Europe/Warsaw'),
    'uk': ('78652228',  'amzn1.vg.6776072 @ A1F83G8C2ARO7P', 'Europe/London'),
    'se': ('108423884', 'amzn1.vg.6968292 @ A2NODRKZP88ZB9', 'Europe/Stockholm'),
}
NEEDED = {'asin', 'orderedUnits', 'orderedRevenue', 'priorOrderedUnits', 'priorOrderedRevenue'}


def mondays():
    """Lunes de las 4 últimas semanas ISO completas (misma regla que build_weekly.py)."""
    today = date.today()
    this_mon = today - timedelta(days=today.weekday())
    return [this_mon - timedelta(weeks=i) for i in range(4, 0, -1)]


def epochs(mon, tz):
    start = datetime(mon.year, mon.month, mon.day, tzinfo=ZoneInfo(tz))
    return int(start.timestamp()), int((start + timedelta(days=7)).timestamp()) - 1


def options(mon, tz):
    f, t = epochs(mon, tz)
    pf, pt = epochs(mon - timedelta(weeks=52), tz)
    return {'fromDate': f, 'toDate': t, 'priorFromDate': pf, 'priorToDate': pt, 'timezone': tz}


def yoy_key(mon):
    y, w, _ = (mon - timedelta(weeks=52)).isocalendar()
    return f'{y}_W{w:02d}'


def plan():
    mons = mondays()
    out = {'weeks': [{'slot': i, 'iso': f'{m.isocalendar()[0]}-W{m.isocalendar()[1]:02d}', 'monday': m.isoformat()}
                     for i, m in enumerate(mons)], 'markets': {}}
    for code, (ch, mer, tz) in MARKETS.items():
        missing = []
        for m in mons:
            key = yoy_key(m)
            if not os.path.exists(os.path.join(DATA, 'yoy', code, f'{key}.csv')):
                prev_mon = m - timedelta(weeks=52)
                missing.append({'key': key, 'reportOptions': options(prev_mon, tz)})
        out['markets'][code] = {'channelId': ch, 'merchantId': mer,
                                'weeks': [{'slot': i, 'reportOptions': options(m, tz)} for i, m in enumerate(mons)],
                                'yoy_missing': missing}
    print(json.dumps(out, indent=1))


def week_range(mon):
    return f'{mon:%Y_%m_%d}-{mon + timedelta(days=6):%Y_%m_%d}'


def fetch(url, dst, mon):
    if not url:  # informe "errored": sin datos esa semana -> CSV solo con cabecera
        with open(dst, 'w', encoding='utf-8', newline='') as f:
            csv.writer(f).writerow(sorted(NEEDED))
        return 0, 0.0, 'sin datos'
    if week_range(mon) not in url:
        raise ValueError(f'el informe no es lunes-domingo {week_range(mon)}: {url.split("?")[0][-80:]}')
    urllib.request.urlretrieve(url, dst)
    with open(dst, encoding='utf-8-sig') as f:
        rd = csv.DictReader(f)
        if not NEEDED <= set(rd.fieldnames or []):
            raise ValueError(f'cabecera inesperada en {dst}: {rd.fieldnames}')
        u = r = 0
        for row in rd:
            u += float(row['orderedUnits'] or 0); r += float(row['orderedRevenue'] or 0)
    return u, r, ''


def market_dir(code):
    return DATA if code == 'es' else os.path.join(DATA, code)


def units(path):
    if not os.path.exists(path):
        return None
    with open(path, encoding='utf-8-sig') as f:
        return sum(float(row.get('orderedUnits') or 0) for row in csv.DictReader(f))


def download(urls_path):
    urls = json.load(open(urls_path, encoding='utf-8'))
    mons = mondays()
    tmp = tempfile.mkdtemp(prefix='finocam_weekly_')
    moves, errors = [], []
    for code in MARKETS:
        mu = urls.get(code)
        if mu is None:
            errors.append(f'{code}: falta en urls.json'); continue
        for i, mon in enumerate(mons):
            dst = os.path.join(tmp, f'{code}_weekly_{i}.csv')
            try:
                u, r, note = fetch(mu.get(str(i)), dst, mon)
            except Exception as e:
                errors.append(f'{code} semana {i}: {e}'); continue
            old = units(os.path.join(market_dir(code), f'weekly_{i + 1}.csv')) if i < 3 else None
            print(f'{code} {mon.isocalendar()[0]}-W{mon.isocalendar()[1]:02d}: {u:>7.0f} uds {r:>11,.0f} '
                  f'| antes {"-" if old is None else f"{old:.0f}"} {note}')
            moves.append((dst, os.path.join(market_dir(code), f'weekly_{i}.csv')))
        for key, url in (mu.get('yoy') or {}).items():
            y, w = int(key[:4]), int(key[6:])
            dst = os.path.join(tmp, f'{code}_{key}.csv')
            try:
                u, r, note = fetch(url, dst, date.fromisocalendar(y, w, 1))
            except Exception as e:
                errors.append(f'{code} yoy {key}: {e}'); continue
            print(f'{code} {key} (año anterior): {u:>7.0f} uds {r:>11,.0f} {note}')
            os.makedirs(os.path.join(DATA, 'yoy', code), exist_ok=True)
            moves.append((dst, os.path.join(DATA, 'yoy', code, f'{key}.csv')))
    if errors:
        print('\nERRORES — no se ha copiado nada:\n  ' + '\n  '.join(errors))
        sys.exit(1)
    for src, dst in moves:
        shutil.copyfile(src, dst)
    shutil.rmtree(tmp, ignore_errors=True)
    print(f'\nOK: {len(moves)} ficheros copiados a {DATA}')


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    if len(sys.argv) >= 2 and sys.argv[1] == 'plan':
        plan()
    elif len(sys.argv) == 3 and sys.argv[1] == 'download':
        download(sys.argv[2])
    else:
        sys.exit(__doc__)
