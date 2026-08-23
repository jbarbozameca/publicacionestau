#!/usr/bin/env python3
"""Ensambla el portal: plantilla + datos + logo -> portal_tau.html"""
import json, base64, sys, subprocess, pathlib, datetime

BASE = pathlib.Path('/home/claude')
xlsx = sys.argv[1] if len(sys.argv) > 1 else None

if xlsx:
    subprocess.run([sys.executable, str(BASE/'build_data.py'), xlsx, str(BASE/'data.json')], check=True)

data = json.load(open(BASE/'data.json', encoding='utf-8'))
data['actualizado'] = datetime.date.today().isoformat()
tpl = open(BASE/'portal_template.html', encoding='utf-8').read()
logo = base64.b64encode(open(BASE/'mark.png','rb').read()).decode()

html = tpl.replace('/*__DATA__*/', json.dumps(data, ensure_ascii=False)).replace('__LOGO__', logo)
out = BASE/'portal_tau.html'
out.write_text(html, encoding='utf-8')
print('escrito', out, round(len(html)/1024), 'KB')
