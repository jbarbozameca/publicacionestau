#!/usr/bin/env python3
"""Convierte el Excel de transparencia Tau en el JSON que consume el portal."""
import openpyxl, json, re, sys, unicodedata, datetime

if len(sys.argv) < 2:
    sys.exit('Uso: python3 build_data.py <archivo.xlsx> [salida.json]')
XLSX = sys.argv[1]
OUT = sys.argv[2] if len(sys.argv) > 2 else 'data.json'

MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']

STATUS_MAP = {
    'PUBLICADO': 'Publicado',
    'ACEPTADO': 'Aceptado',
    'PEER-REVIEW (R1)': 'En peer review (R1)',
    'PEER REVIEW (R1)': 'En peer review (R1)',
    'PEER REVIEW': 'En peer review (R1)',
    'ENVIADO': 'Enviado',
    'EN REDACCIÓN': 'En redacción',
    'EN REDACCION': 'En redacción',
    'EN PROCESO': 'En proceso',
    'RECHAZADO': 'Rechazado',
}

NOISE = re.compile(r'(universidad|university|facultad|school of|department|departamento|hospital|instituto|institute|escuela|centro de|research center|peru|perú|méxico|mexico|colombia|ecuador|chile|españa|spain|argentina|brasil)', re.I)


def clean_cell(v):
    if v is None:
        return ''
    if isinstance(v, datetime.datetime):
        return v.strftime('%Y-%m-%d')
    s = str(v).strip()
    if s.endswith('.0') and re.fullmatch(r'\d+\.0', s):
        s = s[:-2]
    return re.sub(r'[ \t]+', ' ', s).strip()


def norm_status(s):
    if not s:
        return 'Sin estado'
    key = re.sub(r'\s+', ' ', s.upper().strip())
    if key in STATUS_MAP:
        return STATUS_MAP[key]
    for k, v in STATUS_MAP.items():
        if key.startswith(k):
            return v
    return s.title()


def group_label(raw, sheet):
    """Devuelve la etiqueta de grupo o None si la celda no identifica un grupo."""
    if not raw:
        return None
    s = raw.strip()
    m = re.fullmatch(r'(\d{4})-(\d{2})-(\d{2})', s)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        base = 'Training' if sheet == 'TRAINING' else 'Grupo'
        return f'{base} {MESES[mo-1].capitalize()} {y}'
    if re.fullmatch(r'\d+(\.\d+)?', s):        # solo un correlativo -> no es grupo
        return None
    if s.lower() in ('nº', 'no', 'nº'):
        return None
    # normaliza mayúsculas iniciales de "focus enero 2026" -> "Focus Enero 2026"
    words = s.split()
    fixed = []
    for w in words:
        if re.fullmatch(r'\d{4}', w) or w.isupper():
            fixed.append(w)
        else:
            fixed.append(w[:1].upper() + w[1:])
    return ' '.join(fixed)


def strip_marks(tok):
    tok = tok.replace('*', ' ')
    tok = re.sub(r'[¹²³⁰-⁹ⁱⁿ⁻⁺]+', ' ', tok)   # superíndices
    tok = re.sub(r'[,;]?\s*\b(PhD|MD|MSc|MPH|Dr\.?|Dra\.?|Mg\.?|Esp\.?)\b\.?', ' ', tok, flags=re.I)
    tok = re.sub(r'(?<=[A-Za-záéíóúñÁÉÍÓÚÑ\.\)])\s*[\d,\-– ]+$', '', tok)  # afiliaciones finales
    tok = re.sub(r'\d', ' ', tok)
    tok = re.sub(r'\s+', ' ', tok).strip(' .,-;:()')
    return tok


def parse_authors(raw):
    if not raw:
        return []
    txt = raw.replace('\r', '\n')
    txt = re.sub(r'\n+', ';', txt)
    txt = re.sub(r'\s+(and|y)\s+', ';', txt)
    parts = []
    for chunk in txt.split(';'):
        parts.extend(re.split(r',(?![^(]*\))', chunk))
    out = []
    for p in parts:
        p = strip_marks(p)
        if not p or len(p) < 4:
            continue
        if len(p.split()) < 2:          # un solo token: iniciales sueltas o restos
            continue
        if NOISE.search(p):
            continue
        if len(p.split()) > 6:          # frases, no nombres
            continue
        if p.lower().startswith(('efficacy', 'effect', 'assessment', 'prognostic', 'comparison')):
            continue
        if p not in out:
            out.append(p)
    return out


def slug(s):
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return re.sub(r'[^a-z0-9]+', ' ', s.lower()).strip()


def main():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    records = []
    rid = 0
    sheet_labels = {
        'ACEPTADOS PUBLICADOS': 'Aceptados / Publicados',
        'TRAINING': 'Training',
        'FOCUS': 'Focus',
    }
    for ws in wb.worksheets:
        if ws.title not in sheet_labels:
            continue
        for i, row in enumerate(ws.iter_rows(min_row=2, max_col=8, values_only=True), 2):
            v = [clean_cell(c) for c in row]
            while len(v) < 8:
                v.append('')
            titulo = v[1]
            if not titulo or titulo.lower().startswith('título de trabajo'):
                continue
            autores_raw = v[2]
            # celdas donde el título se repitió en la columna de autores
            if autores_raw and slug(autores_raw)[:40] == slug(titulo)[:40]:
                autores_raw = ''
            g = group_label(v[0], ws.title)
            rid += 1
            records.append({
                'id': rid,
                'seccion': sheet_labels[ws.title],
                'grupo': g or 'Sin grupo asignado',
                'titulo': titulo,
                'autoresRaw': autores_raw,
                'autores': parse_authors(autores_raw),
                'estado': norm_status(v[3]),
                'revista': v[4] if v[4] and v[4].upper() != 'PENDIENTE DE ASIGNACIÓN' else '',
                'cuartil': v[5],
                'link': v[6] if v[6].lower().startswith('http') else '',
                'obs': re.sub(r'\s*\n\s*', ' ', v[7]).strip(),
            })

    # celdas de autores que en realidad contienen el título de otro trabajo
    titulos = {slug(r['titulo'])[:40] for r in records}
    for r in records:
        if r['autoresRaw'] and slug(r['autoresRaw'])[:40] in titulos:
            r['autoresRaw'] = ''
            r['autores'] = []

    # índice de autores
    authors = {}
    for r in records:
        for a in r['autores']:
            k = slug(a)
            if not k:
                continue
            authors.setdefault(k, {'nombre': a, 'n': 0, 'variantes': set()})
            authors[k]['n'] += 1
            authors[k]['variantes'].add(a)
    idx = sorted(
        ({'nombre': sorted(v['variantes'], key=len)[-1], 'key': k, 'n': v['n']} for k, v in authors.items()),
        key=lambda x: (-x['n'], x['nombre'])
    )

    data = {
        'actualizado': datetime.date.today().isoformat(),
        'registros': records,
        'autores': idx,
    }
    json.dump(data, open(OUT, 'w'), ensure_ascii=False)
    print('registros:', len(records), '| autores únicos:', len(idx))
    from collections import Counter
    print('estados:', Counter(r['estado'] for r in records).most_common())
    print('grupos:', Counter(r['grupo'] for r in records).most_common())


if __name__ == '__main__':
    main()
