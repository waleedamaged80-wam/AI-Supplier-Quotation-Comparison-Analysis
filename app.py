"""Supplier quotation comparison - Streamlit app.   Run:  streamlit run app.py"""
import hashlib
import os
import shutil

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

import engine
import rules

st.set_page_config(page_title="RFQ Compliance Comparison", page_icon="📊", layout="wide")


def secret(name):
    try:
        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:
        pass
    return os.environ.get(name)


# ---------- password gate ----------
pw = secret("APP_PASSWORD")
if pw and not st.session_state.get("authed"):
    st.title("RFQ Compliance Comparison")
    entered = st.text_input("Password", type="password")
    if entered:
        if entered == pw:
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Wrong password")
    st.stop()

ss = st.session_state
ss.setdefault("docs", None)
ss.setdefault("rows", None)
ss.setdefault("decisions", {})
ss.setdefault("rates", {})
ss.setdefault("adj", {})
ss.setdefault("errors", [])
ss.setdefault("xlsx", None)
ss.setdefault("baseline", None)

# ---------- sidebar ----------
with st.sidebar:
    st.header("Settings")
    has_cli = shutil.which("claude") is not None
    opts = (["claude-code"] if has_cli else []) + ["api"]
    labels = {"claude-code": "Claude subscription (Claude Code, this PC only)", "api": "Anthropic API key (pay per use)"}
    backend = st.radio("Extraction engine", opts, format_func=labels.get)
    api_key = None
    if backend == "api":
        api_key = secret("ANTHROPIC_API_KEY")
        if not api_key:
            api_key = st.text_input("Anthropic API key", type="password") or None
        else:
            st.caption("API key loaded from secrets.")
    elif not has_cli:
        st.caption("Claude Code not found on this machine.")
    threshold = st.slider("Review threshold (AI confidence)", 0.5, 1.0, 0.8, 0.05,
                          help="Extracted values below this confidence are shown YELLOW until you tick Reviewed.")
    st.caption("Up to 8 suppliers, 3 BOQ items, 30 requirement lines. Quotation PDFs are sent to Claude for reading.")

st.title("Supplier quotation comparison")
st.caption("Compare like-for-like before comparing price. AI reads, rules decide, exceptions need a named approver.")

# ---------- 1. baseline ----------
st.subheader("1. RFQ baseline")
tpl = rules.load_baseline(load_workbook(engine.TEMPLATE))
tboq, tlines = engine.baseline_tables(tpl)
empty = st.checkbox("Start with an empty baseline (new RFQ)")
c1, c2 = st.columns([1, 3])
rfq_id = c1.text_input("RFQ number", "" if empty else str(tpl["header"].get("RFQ ID") or ""))
title = c2.text_input("Title", "" if empty else str(tpl["header"].get("Title") or ""))
k = "empty" if empty else "tpl"
boq_df = pd.DataFrame([dict(Item=1, Description="", Unit="", Qty=0)] if empty else tboq)
lines_df = pd.DataFrame(
    [dict(Line=1, Parameter="", Class="Mandatory", Rule="Exact", Target="", Tolerance=None, **{"Fail result": "RED"}, Notes="")]
    if empty else tlines)
st.markdown("**BOQ items** (priced items, max 3)")
boq_ed = st.data_editor(boq_df, num_rows="dynamic", width="stretch", hide_index=True, key=f"boq_{k}")
st.markdown("**Requirement lines** (what every supplier must meet)")
lines_ed = st.data_editor(
    lines_df, num_rows="dynamic", width="stretch", hide_index=True, key=f"lines_{k}",
    column_config={
        "Line": st.column_config.NumberColumn(step=1),
        "Class": st.column_config.SelectboxColumn(options=engine.CLASSES, required=True),
        "Rule": st.column_config.SelectboxColumn(options=engine.RULES, required=True),
        "Fail result": st.column_config.SelectboxColumn(options=engine.FAILS, required=True),
        "Target": st.column_config.TextColumn(help="Number, text, or date as YYYY-MM-DD"),
    })
baseline = engine.make_baseline(rfq_id, title, boq_ed.to_dict("records"), lines_ed.to_dict("records"), threshold)
errs = engine.validate(baseline)
for e in errs:
    st.error(e)

# ---------- 2. uploads ----------
st.subheader("2. Quotations")
files = st.file_uploader(f"Upload up to {engine.MAX_SUPPLIERS} quotation PDFs (one PDF per supplier)",
                         type="pdf", accept_multiple_files=True)
if len(files) > engine.MAX_SUPPLIERS:
    st.error(f"Maximum {engine.MAX_SUPPLIERS} quotations.")
    files = files[:engine.MAX_SUPPLIERS]
names = []
if files:
    cols = st.columns(min(4, len(files)))
    for i, f in enumerate(files):
        default = os.path.splitext(f.name)[0].replace("_", " ").replace("-", " ").title()
        names.append(cols[i % len(cols)].text_input(f"Supplier {i + 1}", default, key=f"nm_{i}_{f.name}").strip())
    if len(set(names)) != len(names) or "" in names:
        st.error("Supplier names must be unique and not empty.")
        errs = errs + ["names"]

# ---------- 3. run ----------
st.subheader("3. Run")
need_key = backend == "api" and not api_key
if need_key:
    st.warning("Enter an Anthropic API key in the sidebar (or choose the Claude subscription option).")
go = st.button("Run comparison", type="primary", disabled=bool(errs) or not files or need_key)
if go:
    bar = st.progress(0.0, text="Reading quotations...")
    payload = [(n, f.name, f.getvalue()) for n, f in zip(names, files)]

    def prog(done, total, who):
        bar.progress(done / total, text=f"Finished {who} ({done}/{total})")

    docs, errors = engine.extract_all(payload, baseline, backend, api_key, prog)
    bar.empty()
    ss.docs, ss.errors = docs, errors
    ss.rows = engine.make_rows(docs, baseline) if docs else None
    ss.baseline = baseline
    ss.decisions, ss.xlsx = {}, None
    ss.rates = {d["supplier"]: {r["item"]: r["rate"] for r in (d.get("rates") or [])} for d in docs}
    ss.adj = {d["supplier"]: 0 for d in docs}
for s, msg in ss.errors:
    st.error(f"{s}: {msg}")

if not ss.docs:
    st.info("Upload quotations and press Run comparison.")
    st.stop()

baseline = ss.baseline
baseline["threshold"] = threshold
suppliers = [d["supplier"] for d in ss.docs]
results = st.container()

# ---------- 4. review extracted values ----------
st.subheader("4. Check extracted values")
st.caption("Correct anything the AI got wrong. Tick Reviewed when you have checked a value against the PDF (counts as 100% confidence).")
df = pd.DataFrame(ss.rows, columns=engine.ROW_COLS)
df["_low"] = (df["Confidence"] < threshold) & ~df["Reviewed"]
order = df.sort_values(["_low", "Supplier", "Line"], ascending=[False, True, True], kind="stable").drop(columns="_low")
show_low = st.checkbox("Show only low-confidence / unreviewed-below-threshold values")
view = order[order["Confidence"] < threshold] if show_low else order
edited = st.data_editor(
    view, hide_index=True, width="stretch", key=f"rows_{show_low}_{len(ss.docs)}",
    disabled=["Supplier", "Line", "Parameter", "Confidence", "Source"],
    column_config={
        "Declaration": st.column_config.SelectboxColumn(options=engine.DECLARATIONS, required=True),
        "Confidence": st.column_config.ProgressColumn(min_value=0, max_value=1, format="%.2f"),
        "Reviewed": st.column_config.CheckboxColumn(),
        "Source": st.column_config.TextColumn(width="large"),
    })
upd = {(r["Supplier"], int(r["Line"])): r for r in edited.to_dict("records")}
ss.rows = [upd.get((r["Supplier"], int(r["Line"])), r) for r in ss.rows]

res_open, base_res = engine.compute(ss.rows, baseline)

# ---------- 5. exceptions ----------
st.subheader("5. Exceptions")
items = engine.open_items(base_res, ss.rows, baseline)
if not items:
    st.success("No open RED/YELLOW items.")
else:
    st.caption("A line is cleared only when Decision = Accepted AND an approver is named. Claude never accepts an exception.")
    ex_df = pd.DataFrame([dict(**i, Decision=ss.decisions.get((i["Supplier"], i["Line"]), {}).get("Decision"),
                               Approver=ss.decisions.get((i["Supplier"], i["Line"]), {}).get("Approver", ""),
                               Justification=ss.decisions.get((i["Supplier"], i["Line"]), {}).get("Justification", ""))
                          for i in items])
    sig = hashlib.md5("|".join(f"{i['Supplier']}{i['Line']}{i['Status']}" for i in items).encode()).hexdigest()[:8]
    ex_ed = st.data_editor(
        ex_df, hide_index=True, width="stretch", key=f"ex_{sig}",
        disabled=["Supplier", "Line", "Parameter", "Status", "Required", "Offered", "Note"],
        column_config={"Decision": st.column_config.SelectboxColumn(options=engine.DECISIONS)})
    dec = {}
    for r in ex_ed.to_dict("records"):
        d = engine.clean(r.get("Decision"))
        if d:
            dec[(r["Supplier"], int(r["Line"]))] = dict(Decision=d, Approver=engine.clean(r.get("Approver")) or "",
                                                         Justification=engine.clean(r.get("Justification")) or "",
                                                         Offered=r.get("Offered", ""))
    ss.decisions = dec

# ---------- 6. price ----------
st.subheader("6. Rates and adjustments")
st.caption("Rates are read from the quotation where possible. Adjustments add cost for scope gaps (e.g. missing item, freight).")
rd = pd.DataFrame([dict(Supplier=s, **{f"Item {b['item']} rate": ss.rates.get(s, {}).get(b["item"]) for b in baseline["boq"]},
                        Adjustment=ss.adj.get(s, 0)) for s in suppliers])
rd_ed = st.data_editor(rd, hide_index=True, width="stretch", disabled=["Supplier"], key=f"rates_{len(suppliers)}")
for r in rd_ed.to_dict("records"):
    s = r["Supplier"]
    ss.rates[s] = {b["item"]: engine.clean(r.get(f"Item {b['item']} rate")) for b in baseline["boq"]}
    ss.adj[s] = engine.clean(r.get("Adjustment")) or 0

# ---------- results (top of page) ----------
accepted = engine.accepted_set(ss.decisions)
final, base_res = engine.compute(ss.rows, baseline, accepted)
ev = engine.evaluate(suppliers, ss.rates, ss.adj, baseline, final)

COL = {"GREEN": "#c6efce", "YELLOW": "#ffeb9c", "RED": "#ffc7ce", "CLEARED": "#bdd7ee"}
with results:
    st.subheader("Results")
    summ = pd.DataFrame([dict(Supplier=s, Overall=final[s]["overall"], Eligible=final[s]["eligible"], RED=final[s]["red"],
                              YELLOW=final[s]["yellow"], Cleared=final[s]["cleared"]) for s in suppliers])
    st.dataframe(summ, hide_index=True, width="stretch")
    mat = pd.DataFrame({s: {f"{n}. {baseline['lines'][n]['parameter']}":
                            final[s]["statuses"][n]
                            for n in baseline["lines"]} for s in suppliers})
    st.markdown("**Compliance matrix**")
    st.dataframe(mat.style.map(lambda v: f"background-color:{COL.get(v, '')}"), width="stretch")
    st.markdown("**Price evaluation** (only eligible suppliers with a complete price are ranked)")
    st.dataframe(pd.DataFrame(ev), hide_index=True, width="stretch",
                 column_config={c: st.column_config.NumberColumn(format="%,.0f") for c in ("Quoted", "Adjustments", "Indicative", "Evaluated")})
    st.caption("Ranking is indicative. The award decision stays with Procurement.")
    if st.button("Build Excel file"):
        with st.spinner("Building workbook..."):
            ss.xlsx = engine.build_workbook(baseline, ss.docs, ss.rows, ss.rates, ss.adj, ss.decisions)
    if ss.xlsx:
        st.download_button("Download Excel workbook", ss.xlsx,
                           file_name=f"RFQ_{baseline['header'].get('RFQ ID') or 'comparison'}_comparison.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("Rebuild after any later edit.")
