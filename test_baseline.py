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
rid, ttl, boq, lines = be.extract_api(C, 'rfq.txt', b'Doors 10 m2, min 1.5 mm')
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
r = be.extract_baseline('rfq.pdf', b'%PDF-1.4 fake', 'claude-code')
assert r[0] == 'P09999' and len(r[3]) == 3
print('OK')
