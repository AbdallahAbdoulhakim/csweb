#!/usr/bin/env python3
"""
Génère le SQL de chargement du cadre cartographique (base ucrgph05_cadre)
à partir du classeur du cadre (ex. CARTO_PILOTE_AVT.xlsx).

Feuilles lues : PreZD, ILOTS, STR. Les niveaux supérieurs (île, préfecture,
commune, village) se déduisent des PreZD.

Codes cumulatifs (mêmes formats que les zones du dashboard) :
  ile 1 | prefecture 3 | commune 4 | village 7 (village 2 + UR 1)
  prezd 10 | ilot 14 | structure 19

Usage (aucune dépendance externe) :
  python3 -I charger_cadre.py CARTO_PILOTE_AVT.xlsx > cadre.sql
  docker exec -i csweb-mysql-1 sh -c 'mysql -uroot -p"$MYSQL_ROOT_PASSWORD"' < cadre.sql

Le script recrée les tables : il peut être relancé à chaque nouvelle version du cadre.
"""
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
DB = 'ucrgph05_cadre'


def read_sheets(path, wanted):
    z = zipfile.ZipFile(path)
    strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        for si in ET.fromstring(z.read('xl/sharedStrings.xml')).iter(NS + 'si'):
            strings.append(''.join(t.text or '' for t in si.iter(NS + 't')))
    rels = {r.get('Id'): r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    sheets = {}
    for s in ET.fromstring(z.read('xl/workbook.xml')).iter(NS + 'sheet'):
        if s.get('name') not in wanted:
            continue
        target = rels[s.get(REL + 'id')].lstrip('/')
        target = target if target.startswith('xl/') else 'xl/' + target
        rows = []
        for row in ET.fromstring(z.read(target)).iter(NS + 'row'):
            cells = {}
            for c in row.iter(NS + 'c'):
                v = c.find(NS + 'v')
                val = v.text if v is not None else None
                if c.get('t') == 's' and val is not None:
                    val = strings[int(val)]
                elif c.get('t') == 'inlineStr':
                    val = ''.join(t.text or '' for t in c.iter(NS + 't'))
                cells[re.sub(r'\d', '', c.get('r'))] = val
            rows.append(cells)
        header = rows[0]
        sheets[s.get('name')] = [
            {header[col]: r.get(col) for col in header} for r in rows[1:] if any(r.values())
        ]
    missing = set(wanted) - set(sheets)
    if missing:
        sys.exit(f'Feuilles absentes : {missing}')
    return sheets


def code(value, length):
    """Code numérique en texte complété de zéros (Excel peut perdre les zéros de tête)."""
    if value is None or str(value).strip() == '':
        return None
    v = str(value).strip()
    if re.fullmatch(r'\d+\.0+', v):
        v = v.split('.')[0]
    if not v.isdigit() or len(v) > length:
        raise ValueError(f'code invalide {value!r} (longueur attendue {length})')
    return v.zfill(length)


def num(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def sql(v):
    if v is None:
        return 'NULL'
    if isinstance(v, (int, float)):
        return repr(v)
    return "'" + str(v).replace('\\', '\\\\').replace("'", "''") + "'"


def geo(r):
    """Colonnes géographiques communes, avec contrôle de cohérence des préfixes."""
    g = {
        'ile': code(r['CodeIle'], 1),
        'prefecture': code(r['NcodePef'], 3),
        'commune': code(r['NcodeCom'], 4),
        'village': code(r['NcodeVVUR'], 7),
        'code_prezd': code(r['CodePreZD'], 10),
        'nom_ile': r.get('ILE'),
        'nom_prefecture': r.get('PREFECTURE'),
        'nom_commune': r.get('COMMUNE'),
        'nom_village': r.get('VILLE_VILLAGE'),
        'ur': int(num(r['UR'])) if num(r.get('UR')) is not None else None,
    }
    chain = [g['ile'], g['prefecture'], g['commune'], g['village'], g['code_prezd']]
    for parent, child in zip(chain, chain[1:]):
        if not child.startswith(parent):
            raise ValueError(f'hiérarchie incohérente {chain}')
    return g


GEO_DDL = """
  ile CHAR(1) NOT NULL, prefecture CHAR(3) NOT NULL, commune CHAR(4) NOT NULL, village CHAR(7) NOT NULL,
  code_prezd CHAR(10) NOT NULL,
  nom_ile VARCHAR(100), nom_prefecture VARCHAR(100), nom_commune VARCHAR(100), nom_village VARCHAR(100),
  ur TINYINT"""
GEO_IDX = """,
  KEY (ile), KEY (prefecture), KEY (commune), KEY (village), KEY (code_prezd)"""

TABLES = {
    'cadre_prezd': ('code_prezd', ['pop', 'x', 'y'],
                    f'CREATE TABLE cadre_prezd ({GEO_DDL}, pop INT, x DOUBLE, y DOUBLE,\n'
                    f'  PRIMARY KEY (code_prezd), KEY (ile), KEY (prefecture), KEY (commune), KEY (village))'),
    'cadre_ilot': ('code_ilot', ['code_ilot', 'pop', 'x', 'y'],
                   f'CREATE TABLE cadre_ilot ({GEO_DDL}, code_ilot CHAR(14) NOT NULL, pop INT, x DOUBLE, y DOUBLE,\n'
                   f'  PRIMARY KEY (code_ilot){GEO_IDX})'),
    'cadre_structure': ('code_structure', ['code_ilot', 'code_structure', 'x', 'y'],
                        f'CREATE TABLE cadre_structure ({GEO_DDL}, code_ilot CHAR(14) NOT NULL,\n'
                        f'  code_structure CHAR(19) NOT NULL, x DOUBLE, y DOUBLE,\n'
                        f'  PRIMARY KEY (code_structure), KEY (code_ilot){GEO_IDX})'),
}


def build_rows(sheets):
    out = {'cadre_prezd': [], 'cadre_ilot': [], 'cadre_structure': []}
    for r in sheets['PreZD']:
        g = geo(r)
        g.update(pop=int(num(r['POP']) or 0), x=num(r['X']), y=num(r['Y']))
        out['cadre_prezd'].append(g)
    for r in sheets['ILOTS']:
        g = geo(r)
        g.update(code_ilot=code(r['CodeILOT'], 14), pop=int(num(r['POP']) or 0), x=num(r['X']), y=num(r['Y']))
        assert g['code_ilot'].startswith(g['code_prezd']), g
        out['cadre_ilot'].append(g)
    for r in sheets['STR']:
        r = dict(r, POP=None)
        g = geo(r)
        g.update(code_ilot=code(r['CodeILOT'], 14), code_structure=code(r['CodeSTR'], 19),
                 x=num(r['X']), y=num(r['Y']))
        assert g['code_structure'].startswith(g['code_ilot']), g
        out['cadre_structure'].append(g)
    return out


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    rows = build_rows(read_sheets(sys.argv[1], {'PreZD', 'ILOTS', 'STR'}))

    print(f'CREATE DATABASE IF NOT EXISTS {DB} CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;')
    print(f'USE {DB};')
    print('SET NAMES utf8mb4;')
    for table, (pk, _, ddl) in TABLES.items():
        data = rows[table]
        keys = [d[pk] for d in data]
        dups = len(keys) - len(set(keys))
        if dups:
            sys.exit(f'{table} : {dups} code(s) en double dans le classeur')
        print(f'DROP TABLE IF EXISTS {table};')
        print(ddl + ' ENGINE=InnoDB;')
        cols = list(data[0].keys())
        for i in range(0, len(data), 1000):
            values = ',\n'.join('(' + ','.join(sql(d[c]) for c in cols) + ')' for d in data[i:i + 1000])
            print(f'INSERT INTO {table} ({",".join(cols)}) VALUES\n{values};')
        print(f'-- {table} : {len(data)} lignes', file=sys.stderr)
    print(f"GRANT SELECT ON {DB}.* TO 'dashboard_ro'@'%';")


if __name__ == '__main__':
    main()
