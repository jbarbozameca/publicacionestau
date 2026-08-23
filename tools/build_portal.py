#!/usr/bin/env python3
"""Ensambla el portal de transparencia: plantilla + datos del Excel + logo.

Uso:
    python3 tools/build_portal.py "data/TRANSPARENCIA - INVESTIGACIONES  Tau 2024-2026.xlsx"

Si no se pasa el Excel, busca el más reciente dentro de ../data/.
Escribe index.html en la raíz del repositorio (y data.json junto al script).
"""
import json, base64, sys, subprocess, pathlib, datetime, glob

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent if HERE.name == 'tools' else HERE
OUT = ROOT / 'index.html'

xlsx = sys.argv[1] if len(sys.argv) > 1 else None
if not xlsx:
    candidatos = sorted(glob.glob(str(ROOT / 'data' / '*.xlsx')), key=lambda p: pathlib.Path(p).stat().st_mtime)
    if not candidatos:
        sys.exit('No encontré ningún .xlsx en data/. Pasa la ruta como argumento.')
    xlsx = candidatos[-1]
    print('Excel:', xlsx)

subprocess.run([sys.executable, str(HERE / 'build_data.py'), xlsx, str(HERE / 'data.json')], check=True)

data = json.load(open(HERE / 'data.json', encoding='utf-8'))
data['actualizado'] = datetime.date.today().isoformat()
tpl = (HERE / 'portal_template.html').read_text(encoding='utf-8')
logo = base64.b64encode((HERE / 'mark.png').read_bytes()).decode()

html = tpl.replace('/*__DATA__*/', json.dumps(data, ensure_ascii=False)).replace('__LOGO__', logo)
OUT.write_text(html, encoding='utf-8')
print('escrito', OUT, round(len(html) / 1024), 'KB ·', len(data['registros']), 'trabajos ·', len(data['autores']), 'autores')
