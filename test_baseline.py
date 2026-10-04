import json, os, stat, sys, tempfile, types, zipfile, io
import baseline_extractor as be, engine

RAW = {'rfq_id': 'P09999', 'title': 'Test RFQ', 'boq': [{'item': 1, 'description': 'Doors', 'unit': 'm2', 'qty': 10}],
       'lines': [{'parameter': 'Item 1 priced', 'cls': 'Mandatory', 'rule': 'Exact', 'target': 'Yes', 'fail': 'RED'},
                 {'parameter': 'Offer valid until', 'cls': 'Evaluated', 'rule': 'Min', 'target': '2026-12-01', 'fail': 'YELLOW'},
                 {'parameter': 'Thickness (mm)', 'cls': 'Mandatory', 'rule': 'Min', 'target': 1.5, 'fail': 'RED'}]}

# API path with a mock client
class B: type, name, input = 'tool_use', 'record_baseline', RAW
class C:
    class messages:
        @staticmethod
        def create(**kw):
            assert kw['tool_choice']['name'] == 'record_baseline'
            return types.SimpleNamespace(content=[B])
rid, ttl, boq, lines = be.extract_api(C, [('doors.txt', b'Doors 10 m2, min 1.5 mm'), ('windows.txt', b'x')])
assert rid == 'P09999' and len(lines) == 3 and lines[2]['Tolerance'] == 0 and lines[2]['Target'] == '1.5'
bl = engine.make_baseline(rid, ttl, boq, lines)
assert not engine.validate(bl), engine.validate(bl)

# docx text extraction
buf = io.BytesIO()
with zipfile.ZipFile(buf, 'w') as z:
    z.writestr('word/document.xml', '<w:document><w:p><w:r><w:t>Min 1.5 mm &amp; ALPHA</w:t></w:r></w:p></w:document>')
assert 'Min 1.5 mm & ALPHA' in be.to_text('x.docx', buf.getvalue())

# Claude Code path with a fake claude
d = tempfile.mkdtemp()
fake = os.path.join(d, 'claude')
open(fake, 'w').write('#!/usr/bin/env python3\nimport json,sys\nprint(json.dumps({"is_error":False,"structured_output":%s}))\n' % json.dumps(RAW))
os.chmod(fake, 0o755)
os.environ['PATH'] = d + os.pathsep + os.environ['PATH']
r = be.extract_baseline([('rfq.pdf', b'%PDF-1.4 fake'), ('b.txt', b'y')], 'claude-code')
assert r[0] == 'P09999' and len(r[3]) == 3
print('OK')

# ---- Gemini path with a mocked call ----
QUOTE = {'supplier_name': 'X', 'lines': [{'line': 1, 'value': 'Yes', 'declaration': 'Fully Compliant', 'confidence': 0.9}], 'rates': []}
import extractor, rules
from openpyxl import load_workbook
calls = []
def fake(key, system, prompt, pdf=None):
    calls.append((key, bool(pdf))); return RAW if 'record_baseline' not in system and 'rfq_id' in system else QUOTE
r = be.extract_gemini('K', [('rfq.pdf', b'%PDF')], call=fake); assert r[0] == 'P09999' and calls[-1] == ('K', True)
r = be.extract_gemini('K', [('rfq.txt', b'text')], call=fake); assert calls[-1] == ('K', False)
bl = rules.load_baseline(load_workbook(engine.TEMPLATE))
import pypdf, io as _io
w = pypdf.PdfWriter(); w.add_blank_page(100, 100); p = os.path.join(tempfile.mkdtemp(), 'q.pdf'); w.write(open(p, 'wb'))
d = extractor.extract_quotation_gemini('K', bl, 'Sup', p, [1, 1], call=fake)
assert d['supplier'] == 'Sup' and len(d['lines']) == len(bl['lines']) and d['lines'][0]['value'] == 'Yes'
print('gemini mock OK')
