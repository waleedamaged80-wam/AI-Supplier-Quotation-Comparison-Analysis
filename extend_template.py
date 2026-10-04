"""One-off: widen P01113_template.xlsx from 8 to 10 suppliers (Compliance Matrix M:N, Evaluation rows 17-18)."""
import copy, sys
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.formatting.formatting import ConditionalFormattingList

src = sys.argv[1] if len(sys.argv) > 1 else 'P01113_template.xlsx'
wb = load_workbook(src)
cm, ev, ap = wb['Compliance Matrix'], wb['Evaluation'], wb['Approval']

def clone(ws, s, d):
    c = ws[d]
    v = ws[s].value
    c.value = Translator(v, origin=s).translate_formula(d) if isinstance(v, str) and v.startswith('=') else v
    c._style = copy.copy(ws[s]._style)

for r in range(1, 43):
    for col in 'MN':
        clone(cm, f'L{r}', f'{col}{r}')
for col in 'FGHIJKL':
    w = cm.column_dimensions[col].width
for col in 'MN':
    cm.column_dimensions[col].width = cm.column_dimensions['E'].width or 13
cm.column_dimensions['M'].width = cm.column_dimensions['N'].width = 13
# keep the title-row style/width the same for E..L as before
old = cm.conditional_formatting
new = ConditionalFormattingList()
for cf in old:
    rng = str(cf.sqref).replace('L36', 'N36').replace('L42', 'N42')
    for rule in cf.rules:
        new.add(rng, rule)
cm.conditional_formatting = new

note = ev['A18'].value
note_style = copy.copy(ev['A18']._style)
ev['A18'].value = None
for r in (17, 18):
    for col in 'ABCDEFGHIJKL':
        clone(ev, f'{col}16', f'{col}{r}')
    ev[f'A{r}'].value = f"=IF('Compliance Matrix'!{'M' if r == 17 else 'N'}4=\"\",\"\",'Compliance Matrix'!{'M' if r == 17 else 'N'}4)"
    ev[f'B{r}'].value = f"=IF(A{r}=\"\",\"\",'Compliance Matrix'!{'M' if r == 17 else 'N'}42)"
ev['A20'].value = note
ev['A20']._style = note_style
for ws in (ev, ap):
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                c.value = c.value.replace('$16', '$18')
# evaluation CF / validation ranges
new = ConditionalFormattingList()
for cf in ev.conditional_formatting:
    rng = ' '.join(x.replace('16', '18') if x.endswith('16') else x for x in str(cf.sqref).split())
    for rule in cf.rules:
        new.add(rng, rule)
ev.conditional_formatting = new
wb.save(src)
print('extended', src)
