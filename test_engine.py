import glob, json, os, subprocess, sys, tempfile
from openpyxl import load_workbook
import engine, rules

EXPECT = {'Tilal Al Eamar': (3, 11), 'Lamset Ebdaa': (5, 6), 'Falcon': (1, 6), 'Mishkal': (2, 6),
          'Qimam': (0, 7), 'Al Shareef': (3, 11), 'Ultra Frame': (0, 6)}
bl = rules.load_baseline(load_workbook(engine.TEMPLATE))
docs = [json.load(open(f, encoding='utf-8')) for f in sorted(glob.glob('fixtures/*.json'))]
names = [d['supplier'] for d in docs]
print(names)
rows = engine.make_rows(docs, bl)
res, base = engine.compute(rows, bl)
for s, r in base.items():
    print(s, r['red'], r['yellow'], r['eligible'])
    exp = EXPECT.get(s)
    assert exp is None or (r['red'], r['yellow']) == exp, (s, exp)
# editing: mark all reviewed -> no low-confidence yellow change on 0.95 fixtures
items = engine.open_items(base, rows, bl)
assert len(items) == sum(r['red'] + r['yellow'] for r in base.values())
dec = {(items[0]['Supplier'], items[0]['Line']): dict(Decision='Accepted', Approver='Test', Justification='x')}
res2, _ = engine.compute(rows, bl, engine.accepted_set(dec))
assert res2[items[0]['Supplier']]['cleared'] >= 1
assert not engine.accepted_set({k: dict(Decision='Accepted', Approver='') for k in dec})
rates = {d['supplier']: {r['item']: r['rate'] for r in d['rates']} for d in docs}
ev = engine.evaluate(names, rates, {}, bl, base)
print([(e['Supplier'], e['Quoted'], e['Rank']) for e in ev])
# baseline round trip
boq, lines = engine.baseline_tables(bl)
b2 = engine.make_baseline('P', 'T', boq, lines)
assert not engine.validate(b2), engine.validate(b2)
assert b2['lines'].keys() == bl['lines'].keys()
for n in bl['lines']:
    assert rules._cmp(rules.fix(b2["lines"][n]["target"])) == rules._cmp(rules.fix(bl["lines"][n]["target"])), n
# workbook export
data = engine.build_workbook(bl, docs, rows, rates, {names[0]: 100}, dec)
p = os.path.join(tempfile.mkdtemp(), 'o.xlsx'); open(p, 'wb').write(data)
subprocess.run([sys.executable, '/mnt/skills/public/xlsx/scripts/recalc.py', p, '120'], check=True)
print(subprocess.run([sys.executable, 'run.py', 'check', '--workbook', p], capture_output=True, text=True).stdout[-1500:])
print('OK')
