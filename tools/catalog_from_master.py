"""Regenera tools/data/catalog_es.json a partir del Màster Famílies de Finocam.

Uso:
    python tools/catalog_from_master.py "ruta/al/MÀSTER FAMÍLIES-2026-2027.xlsx"

Lee la hoja "Refs.Vigents 26-27" (cabecera en la fila 3): ASIN, CATEGORIA (subfamilia),
Anualitat y "descripció ref". La familia sale de la subfamilia; la colección (Moniquilla /
Talkual) de la descripción. Solo entran anualidades 2026*, 2027* y NOCAD, como en build_weekly.py.
"""
import json, os, re, sys
from collections import Counter

import openpyxl

SHEET = 'Refs.Vigents 26-27'
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'catalog_es.json')

# Subfamilia del Màster -> familia del dashboard
SUB2FAM = {
    'Accesorio OP': 'Accesorios',
    'Agenda 16M': 'Agendas', 'Agenda 18M': 'Agendas', 'Agenda Anual': 'Agendas', 'Agenda Curso': 'Agendas',
    'Agenda Docente Curso': 'Agendas', 'Agenda Docente Sin fechar': 'Agendas', 'Agenda Infantil': 'Agendas',
    'Agenda de Anillas Anual': 'Agendas', 'Agenda sin fechar': 'Agendas', 'Reuniones Docente Sin fechar': 'Agendas',
    'Calendario Anual': 'Calendarios', 'Calendario Imán 16M': 'Calendarios', 'Calendario Pared 16M': 'Calendarios',
    'Calendario Pared Anual': 'Calendarios', 'Calendario Pared+Imán Anual': 'Calendarios',
    'Calendario Sobremesa 16M': 'Calendarios', 'Calendario Sobremesa Anual': 'Calendarios',
    'Calendario Vade Anual': 'Calendarios', 'Póster Anual': 'Calendarios',
    'Carpeta': 'Carpetas', 'Carpeta Anillas': 'Carpetas',
    'Cuaderno': 'Cuadernos',
    'Dosier': 'Dosieres',
    'Estuche': 'Estuches',
    'Libro Legalizable': 'Libros de Firma', 'Libro de Firmas': 'Libros de Firma',
    'Planificador sin Fechar': 'Planificadores', 'Planificador Anual': 'Planificadores',
    'Accesorio Portadocumentos': 'Portadocumentos', 'Portadocumentos': 'Portadocumentos',
    'Recambio Portadocumentos': 'Portadocumentos',
    'Recambio Anillas': 'Recambios', 'Recambio Anillas Anual': 'Recambios', 'Recambio Dúo Anual': 'Recambios',
    'Recambio Plana': 'Recambios', 'Recambio Plana Anual': 'Recambios',
    'Índice': 'Índices',
    'Mochila': 'Mochilas', 'Maletín': 'Mochilas',
}


def anualidad(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v or '').strip()


def main(path):
    sys.stdout.reconfigure(encoding='utf-8')
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)[SHEET]
    catalog, unknown, skipped = {}, Counter(), Counter()
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i < 3 or len(r) < 11:
            continue
        asin = str(r[8] or '').strip().upper()
        if not re.fullmatch(r'B0[0-9A-Z]{8}', asin) or asin in catalog:
            continue
        any_ = anualidad(r[4])
        if not (any_.startswith('2026') or any_.startswith('2027') or any_ == 'NOCAD'):
            skipped[any_] += 1
            continue
        sub = str(r[9] or '').strip()
        if sub not in SUB2FAM:
            unknown[sub] += 1
            continue
        desc = str(r[10] or '').strip()
        up = desc.upper()
        col = 'Moniquilla' if 'MONIQUILLA' in up else ('Talkual' if 'TALKUAL' in up else None)
        catalog[asin] = {'familia': SUB2FAM[sub], 'sub': sub, 'any': any_, 'desc': desc, 'col': col}
    if unknown:
        sys.exit(f'Subfamilias sin familia asignada (añádelas a SUB2FAM): {dict(unknown)}')
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(catalog, f, ensure_ascii=False)
    print(f'{len(catalog)} ASIN -> {OUT}')
    print('Por familia:', dict(Counter(v['familia'] for v in catalog.values()).most_common()))
    print('Anualidades descartadas:', dict(skipped))


if __name__ == '__main__':
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
