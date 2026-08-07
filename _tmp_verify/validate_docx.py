"""Throwaway validator: check schema child ordering and column widths in a generated DOCX."""
import sys
import importlib.util
import zipfile
from lxml import etree

spec = importlib.util.spec_from_file_location('xtp', '../xlsform_to_printable.py')
xtp = importlib.util.module_from_spec(spec)
spec.loader.exec_module(xtp)

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
ORDERS = {
    'pPr': xtp._DOCX_PPR_ORDER,
    'rPr': xtp._DOCX_RPR_ORDER,
    'tblPr': xtp._DOCX_TBLPR_ORDER,
    'tcPr': xtp._DOCX_TCPR_ORDER,
    'trPr': xtp._DOCX_TRPR_ORDER,
}

path = sys.argv[1]
with zipfile.ZipFile(path) as zf:
    root = etree.fromstring(zf.read('word/document.xml'))

problems = []
checked = 0
for kind, order in ORDERS.items():
    index = {name.split(':', 1)[1]: pos for pos, name in enumerate(order)}
    for container in root.iter(f'{W}{kind}'):
        checked += 1
        seen = []
        for child in container:
            local = etree.QName(child).localname
            if local not in index:
                problems.append(f'{kind}: unknown child {local}')
                continue
            seen.append((index[local], local))
        positions = [pos for pos, _ in seen]
        if positions != sorted(positions):
            problems.append(f'{kind}: out of order {[name for _, name in seen]}')

print('property containers checked:', checked)
print('ordering problems:', len(problems))
for problem in sorted(set(problems)):
    print('  !', problem)

# Column width consistency across every row of the outer table
outer = root.find(f'.//{W}body/{W}tbl')
grid = [int(c.get(f'{W}w')) for c in outer.find(f'{W}tblGrid')]
print('tblGrid twips:', grid, 'total', sum(grid))
widths = {}
for tr in outer.findall(f'{W}tr'):
    tcs = tr.findall(f'{W}tc')
    row_widths = []
    for tc in tcs:
        tcw = tc.find(f'{W}tcPr/{W}tcW')
        row_widths.append(int(tcw.get(f'{W}w')) if tcw is not None else None)
    widths.setdefault(tuple(row_widths), 0)
    widths[tuple(row_widths)] += 1
for key, count in widths.items():
    print('  row width pattern', key, 'x', count)

nested = outer.findall(f'.//{W}tc/{W}tbl')
print('nested choice tables:', len(nested))
if nested:
    n_grid = [int(c.get(f'{W}w')) for c in nested[0].find(f'{W}tblGrid')]
    n_rows = {tuple(int(tc.find(f'{W}tcPr/{W}tcW').get(f'{W}w')) for tc in tr.findall(f'{W}tc'))
              for tr in nested[0].findall(f'{W}tr')}
    print('  nested grid:', n_grid, 'row patterns:', n_rows)
    for tc in nested[0].findall(f'{W}tr')[0].findall(f'{W}tc'):
        pass

sys.exit(1 if problems else 0)
