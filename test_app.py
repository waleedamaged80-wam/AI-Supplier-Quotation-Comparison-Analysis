import glob, json
from streamlit.testing.v1 import AppTest
import engine, rules
from openpyxl import load_workbook
bl = rules.load_baseline(load_workbook(engine.TEMPLATE))
docs = [json.load(open(f, encoding='utf-8')) for f in sorted(glob.glob('fixtures/*.json'))]
at = AppTest.from_file('app.py', default_timeout=60)
at.session_state['docs'] = docs
at.session_state['baseline'] = bl
at.session_state['rows'] = engine.make_rows(docs, bl)
at.session_state['rates'] = {d['supplier']: {r['item']: r['rate'] for r in d['rates']} for d in docs}
at.session_state['adj'] = {d['supplier']: 0 for d in docs}
at.run()
[r for r in at.radio if 'Baseline' in r.label][0].set_value('P01113 sample').run()
assert not at.exception, at.exception
print('tables', len(at.dataframe), 'editors', len(at.get('arrow_data_frame')))
[b for b in at.button if b.label=='Build Excel file'][0].click().run()   # Build Excel
assert not at.exception, at.exception
print('xlsx bytes', len(at.session_state['xlsx'] or b''))
print('OK')
