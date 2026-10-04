"""Quotation PDF -> structured rows, using the Claude API with a forced tool call.
The AI only extracts and reports confidence. Classification stays in the rules."""
import base64, io, json, os, re, shutil, subprocess, tempfile, time
from datetime import date
from pypdf import PdfReader, PdfWriter

MODEL = os.environ.get('RFQ_MODEL', 'claude-sonnet-5-5')
DECLARATIONS = ['Fully Compliant', 'Partially Compliant', 'Non-Compliant', 'Alternative Offer', 'Not Stated']

SYSTEM = """You extract structured data from supplier quotations (Arabic and English, sometimes scanned or photographed) for a procurement compliance check.
Rules:
- Extract only what the quotation states. If a value is not stated, return null. Never infer, never fill from the RFQ.
- Numeric parameters: return a plain number (no units). Date parameters: ISO date YYYY-MM-DD.
- Text parameters: if the quotation matches the target wording, use the target wording exactly (e.g. "Yes", "Italian", "Wacker GP", "ALPHA", "Ex-VAT"). If it differs, return a short description of what it says (e.g. "German silicone", "EPDM"). Do not force a match.
- Glass total thickness: add the stated layers (glass + spacer + glass). If the quotation states a different total, return the sum as value and say so in note (e.g. "Stated 24 mm; layers sum 22").
- Validity: convert a period (e.g. "one week from 23/09/2026") into the last valid date; give the basis in note.
- "Item N priced": "Yes" if the quotation prices that RFQ item, "No" if it is absent or replaced by a different item, with the reason in note.
- declaration: "Fully Compliant" = the supplier states the value as its offer; "Alternative Offer" = a different brand/spec/equivalent; "Partially Compliant" = partly covered or priced on a different basis; "Non-Compliant" = supplier says it does not offer it; "Not Stated" = absent.
- confidence 0-1: lower it for illegible, handwritten, skewed or ambiguous text and for values you had to interpret. Use <= 0.6 when unsure.
- source_text: the shortest verbatim phrase (original language) supporting the value; source_page: page number inside the PDF you were given.
- rates: unit rate per m2 excluding VAT for each RFQ BOQ item. If only a line total and quantity are given, divide. null if the item is not priced.
Return everything through the record_quotation tool."""


def slice_pdf(path, first, last):
    """Bytes of pages first..last (1-based, inclusive)."""
    reader, writer = PdfReader(path), PdfWriter()
    for i in range(first - 1, last):
        writer.add_page(reader.pages[i])
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


def build_tool():
    return {
        'name': 'record_quotation',
        'description': 'Record the extracted quotation data.',
        'input_schema': {
            'type': 'object',
            'properties': {
                'supplier_name': {'type': 'string'},
                'quotation_date': {'type': ['string', 'null']},
                'lines': {'type': 'array', 'items': {
                    'type': 'object',
                    'properties': {
                        'line': {'type': 'integer'},
                        'value': {'type': ['string', 'number', 'null']},
                        'declaration': {'type': 'string', 'enum': DECLARATIONS},
                        'confidence': {'type': 'number'},
                        'source_page': {'type': ['integer', 'null']},
                        'source_text': {'type': 'string'},
                        'note': {'type': 'string'}},
                    'required': ['line', 'value', 'declaration', 'confidence']}},
                'rates': {'type': 'array', 'items': {
                    'type': 'object',
                    'properties': {'item': {'type': 'integer'}, 'rate': {'type': ['number', 'null']},
                                   'quoted_qty': {'type': ['number', 'null']}},
                    'required': ['item', 'rate']}}},
            'required': ['supplier_name', 'lines', 'rates']}}


def _target_hint(t):
    if hasattr(t, 'isoformat'):
        return f'date, target {t.date().isoformat() if hasattr(t, "date") else t.isoformat()}'
    if isinstance(t, (int, float)):
        return f'number, target {t}'
    return f'text, target "{t}"' if t else 'any'


def build_prompt(baseline, supplier):
    boq = '\n'.join(f"  Item {b['item']}: {b['desc']} - {b['qty']} {b['unit']}" for b in baseline['boq'])
    rows = '\n'.join(
        f"  {n} | {l['parameter']} | {l['rule']} | {_target_hint(l['target'])}"
        for n, l in baseline['lines'].items())
    return (f"RFQ {baseline['header'].get('RFQ ID')} - {baseline['header'].get('Title')}\n"
            f"Quotation from: {supplier}\n\nBOQ items:\n{boq}\n\n"
            f"Requirement lines (line | parameter | rule | type and target). Return one entry per line, in order:\n{rows}\n\n"
            "The attached PDF is this supplier's quotation. Extract it now.")


def build_request(baseline, pdf_bytes, supplier, model=MODEL):
    return dict(
        model=model, max_tokens=8000, system=SYSTEM,
        tools=[build_tool()], tool_choice={'type': 'tool', 'name': 'record_quotation'},
        messages=[{'role': 'user', 'content': [
            {'type': 'document', 'source': {'type': 'base64', 'media_type': 'application/pdf',
                                            'data': base64.standard_b64encode(pdf_bytes).decode()}},
            {'type': 'text', 'text': build_prompt(baseline, supplier)}]}])


def _coerce(value, line):
    """Align the value type with the baseline rule so the rules compare like with like."""
    if value is None or value == '':
        return None
    t = line['target']
    if hasattr(t, 'isoformat'):                        # date target
        try:
            return date.fromisoformat(str(value)[:10])
        except ValueError:
            return str(value)
    if line['rule'] in ('Min', 'Max', 'Range'):        # numeric target
        if isinstance(value, (int, float)):
            return value
        m = re.search(r'-?\d+(?:\.\d+)?', str(value).replace(',', ''))
        return float(m.group()) if m else str(value)
    return str(value)


def normalise(raw, baseline, supplier, pages):
    """Validate the tool output; one row per baseline line, missing ones become Not Stated."""
    by_line = {int(l['line']): l for l in raw.get('lines', [])}
    lines = []
    for n, bl in baseline['lines'].items():
        r = by_line.get(n, {})
        value = _coerce(r.get('value'), bl)
        decl = r.get('declaration') if r.get('declaration') in DECLARATIONS else 'Not Stated'
        if value is None:
            decl = 'Not Stated'
        conf = r.get('confidence')
        conf = 0.0 if conf is None and n not in by_line else (0.5 if conf is None else max(0.0, min(1.0, float(conf))))
        lines.append(dict(line=n, value=value, declaration=decl, confidence=conf,
                          source_page=r.get('source_page'), source_text=(r.get('source_text') or '')[:200],
                          note=(r.get('note') or '')[:200]))
    return dict(supplier=supplier, supplier_name_in_doc=raw.get('supplier_name'),
                quotation_date=raw.get('quotation_date'), pages=pages, model=MODEL,
                rates=[dict(item=int(x['item']), rate=x.get('rate')) for x in raw.get('rates', [])],
                lines=lines)


def extract_quotation(client, baseline, supplier, pdf_path, pages, retries=2):
    pdf = slice_pdf(pdf_path, pages[0], pages[1])
    req = build_request(baseline, pdf, supplier)
    for attempt in range(retries + 1):
        try:
            resp = client.messages.create(**req)
            for block in resp.content:
                if block.type == 'tool_use' and block.name == 'record_quotation':
                    return normalise(block.input, baseline, supplier, pages)
            raise RuntimeError('no tool_use block in response')
        except Exception:
            if attempt == retries:
                raise
            time.sleep(2 * (attempt + 1))


# ---------- Claude Code backend: uses the Claude subscription login, no API credits ----------
CLI_SYSTEM = SYSTEM.replace('Return everything through the record_quotation tool.',
                            'Respond with the structured JSON only.')


def _json_from_text(text):
    """Fallback when structured_output is missing: first {...} block in the reply."""
    m = re.search(r'\{.*\}', text or '', re.S)
    if not m:
        raise RuntimeError('no JSON found in Claude Code reply: ' + (text or '')[:300])
    return json.loads(m.group())


def _cli_env():
    """Drop API credentials so Claude Code uses the subscription login, never API billing."""
    return {k: v for k, v in os.environ.items() if k not in ('ANTHROPIC_API_KEY', 'ANTHROPIC_AUTH_TOKEN')}


def build_cli_command(baseline, supplier, exe='claude'):
    prompt = (CLI_SYSTEM + '\n\n' + build_prompt(baseline, supplier) +
              '\n\nThe quotation is the PDF file quotation.pdf in the current folder. '
              'Use the Read tool to open it and read every page, including scanned pages.')
    schema = build_tool()['input_schema']
    return [exe, '-p', prompt, '--output-format', 'json', '--json-schema', json.dumps(schema),
            '--allowedTools', 'Read']


def extract_with_claude_code(baseline, supplier, pdf_path, pages, retries=1, timeout=900):
    exe = shutil.which('claude')
    if not exe:
        raise RuntimeError('Claude Code is not installed or not on PATH. Run "claude --version" to check.')
    workdir = tempfile.mkdtemp(prefix='rfq_')
    with open(os.path.join(workdir, 'quotation.pdf'), 'wb') as fh:
        fh.write(slice_pdf(pdf_path, pages[0], pages[1]))
    cmd = build_cli_command(baseline, supplier, exe)
    for attempt in range(retries + 1):
        try:
            p = subprocess.run(cmd, cwd=workdir, env=_cli_env(), capture_output=True, text=True,
                               encoding='utf-8', timeout=timeout)
            try:
                resp = json.loads(p.stdout)
            except json.JSONDecodeError:
                raise RuntimeError((p.stdout or p.stderr or 'empty reply')[:500])
            if p.returncode != 0 or resp.get('is_error'):
                raise RuntimeError(str(resp.get('result') or p.stderr)[:500])
            raw = resp.get('structured_output') or _json_from_text(resp.get('result'))
            return normalise(raw, baseline, supplier, pages)
        except Exception:
            if attempt == retries:
                raise
            time.sleep(3)


# ---------- Google Gemini backend: free tier key from aistudio.google.com ----------
GEMINI_MODEL = os.environ.get('RFQ_GEMINI_MODEL', 'gemini-2.5-flash')

_QUOTE_SHAPE = ('Return ONE JSON object only, no markdown, in this shape: {"supplier_name": str, "quotation_date": str|null, '
                '"lines": [{"line": int, "value": str|number|null, "declaration": one of ' + str(DECLARATIONS) +
                ', "confidence": number 0-1, "source_page": int|null, "source_text": str, "note": str}], '
                '"rates": [{"item": int, "rate": number|null}]}')


def gemini_json(api_key, system, prompt, pdf_bytes=None, retries=3):
    """One Gemini call that must return JSON. Retries on rate limits (free tier is slow but free)."""
    from google import genai
    from google.genai import types
    client = genai.Client(api_key=api_key)
    parts = []
    if pdf_bytes is not None:
        parts.append(types.Part.from_bytes(data=pdf_bytes, mime_type='application/pdf'))
    parts.append(prompt)
    cfg = types.GenerateContentConfig(system_instruction=system, response_mime_type='application/json',
                                      temperature=0)
    for attempt in range(retries + 1):
        try:
            resp = client.models.generate_content(model=GEMINI_MODEL, contents=parts, config=cfg)
            return _json_from_text(resp.text)
        except Exception as e:
            msg = str(e)
            if attempt == retries:
                raise
            time.sleep(20 if ('429' in msg or 'RESOURCE_EXHAUSTED' in msg) else 3)


def extract_quotation_gemini(api_key, baseline, supplier, pdf_path, pages, call=None):
    call = call or gemini_json
    system = SYSTEM.replace('Return everything through the record_quotation tool.', _QUOTE_SHAPE)
    raw = call(api_key, system, build_prompt(baseline, supplier), slice_pdf(pdf_path, pages[0], pages[1]))
    return normalise(raw, baseline, supplier, pages)
