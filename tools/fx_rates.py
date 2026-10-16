"""Tipos de cambio del BCE para pasar a € la facturación de UK (GBP), PL (PLN) y SE (SEK).

    python tools/fx_rates.py            actualiza data/fx_eur.json desde el 01/01/2024 hasta hoy

Guarda el tipo de referencia diario del BCE (unidades de moneda por 1 €). build_weekly.py convierte
cada importe con el tipo de su día (media de los días del periodo: semana, mes o acumulado). Los días
sin tipo (fines de semana y festivos) usan el último tipo publicado antes; el histórico entero se
convierte igual, así que una semana de 2025 va con los tipos de 2025.
"""
import csv, io, json, os, sys, urllib.request
from datetime import date, timedelta

DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
FX_JSON = os.path.join(DATA, 'fx_eur.json')
CURRENCIES = ('GBP', 'PLN', 'SEK')
START = '2024-01-01'
URL = ('https://data-api.ecb.europa.eu/service/data/EXR/D.{curs}.EUR.SP00.A'
       '?startPeriod={start}&format=csvdata')


def update():
    url = URL.format(curs='+'.join(CURRENCIES), start=START)
    with urllib.request.urlopen(url, timeout=60) as resp:
        text = resp.read().decode('utf-8')
    rates = {c: {} for c in CURRENCIES}
    for row in csv.DictReader(io.StringIO(text)):
        if row['CURRENCY'] in rates and row['OBS_VALUE']:
            rates[row['CURRENCY']][row['TIME_PERIOD']] = float(row['OBS_VALUE'])
    for c in CURRENCIES:
        if len(rates[c]) < 200:
            sys.exit(f'El BCE devolvió muy pocos tipos para {c} ({len(rates[c])}); no se guarda nada')
    rates = {c: dict(sorted(v.items())) for c, v in rates.items()}
    with open(FX_JSON, 'w', encoding='utf-8') as f:
        json.dump({'source': 'BCE, tipo de referencia diario (moneda por 1 EUR)', 'updated': date.today().isoformat(),
                   'rates': rates}, f, ensure_ascii=False)
    for c in CURRENCIES:
        last = max(rates[c])
        print(f'{c}: {len(rates[c])} días, último {last} = {rates[c][last]}')


class FX:
    """Conversión a €: FX().to_eur(importe, 'GBP', desde, hasta) con la media diaria del periodo."""

    def __init__(self, path=FX_JSON):
        with open(path, encoding='utf-8') as f:
            self.rates = json.load(f)['rates']
        self._daily = {}

    def day_rate(self, cur, d):
        """Tipo del día d (date); si ese día no hay publicación, el último anterior."""
        key = (cur, d)
        if key not in self._daily:
            table = self.rates[cur]
            # último tipo publicado hasta ese día; si no hay (inicio de la serie), el siguiente publicado
            for probe in [d - timedelta(days=i) for i in range(15)] + [d + timedelta(days=i) for i in range(1, 15)]:
                v = table.get(probe.isoformat())
                if v:
                    break
            else:
                raise KeyError(f'Sin tipo de cambio {cur} para {d}: ejecuta python tools/fx_rates.py')
            self._daily[key] = v
        return self._daily[key]

    def avg_rate(self, cur, d1, d2):
        n, s, d = 0, 0.0, d1
        while d <= d2:
            s += self.day_rate(cur, d); n += 1; d += timedelta(days=1)
        return s / n

    def to_eur(self, amount, cur, d1, d2=None):
        if not cur or cur == 'EUR':
            return amount
        return amount / self.avg_rate(cur, d1, d2 or d1)


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    update()
