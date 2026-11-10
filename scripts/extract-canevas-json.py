"""Extrait du canevas la structure de chaque onglet de saisie (consignes, cas, erreurs, colonnes) en JSON pour le guide."""
import json, sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1])
out = {}
for ws in wb.worksheets:
    if not ws.title[0].isdigit():
        continue
    z = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str):
            for k in ('COMMENT REMPLIR', 'CAS POSSIBLES', 'ERREURS', 'VOTRE SAISIE'):
                if v.startswith(k): z[k] = r
    nc = ws.max_column
    hdr = z['VOTRE SAISIE'] + 1
    cols = [{'nom': ws.cell(hdr, c).value, 'aide': ws.cell(hdr + 1, c).value} for c in range(2, nc + 1)]
    steps = [ws.cell(r, 1).value for r in range(z['COMMENT REMPLIR'] + 1, z['CAS POSSIBLES'] - 1) if ws.cell(r, 1).value]
    cases, cur = [], None
    for r in range(z['CAS POSSIBLES'] + 2, z['ERREURS'] - 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith('Cas '):
            cur = {'titre': v, 'exemples': []}; cases.append(cur)
        elif v == 'Ex.' and cur is not None:
            cur['exemples'].append({ws.cell(z['CAS POSSIBLES'] + 1, c).value: ws.cell(r, c).value for c in range(2, nc) if ws.cell(r, c).value not in (None, '')})
    errors = [ws.cell(r, 1).value.lstrip('✗ ').strip() for r in range(z['ERREURS'] + 1, z['VOTRE SAISIE'] - 1) if ws.cell(r, 1).value]
    out[ws.title] = {'titre': ws.cell(1, 1).value, 'bandeau': ws.cell(2, 1).value, 'consignes': steps, 'cas': cases, 'erreurs': errors, 'colonnes': cols}
json.dump(out, open(sys.argv[2], 'w'), ensure_ascii=False, indent=1)
print(list(out))
