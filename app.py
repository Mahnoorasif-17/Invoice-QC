"""
Invoice QC Tool — Streamlit version.

Local-only — no API keys, no network calls at runtime.
  Excel parsing       → pandas + openpyxl
  PDF data extraction → pdfplumber
  QC rules            → qc_engine.run_qc()

Run:    streamlit run app.py
"""

from collections import defaultdict
from datetime import datetime

import pandas as pd
import streamlit as st

from qc_engine import build_context, run_qc, auto_detect_month, MF
from pdf_extractor import extract_pdf


# ─── PAGE CONFIG ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Invoice QC Tool",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
  .main .block-container { padding-top: 1.2rem; padding-bottom: 3rem; max-width: 1200px; }
  .qc-hdr {
    background:#1a2942;color:#fff;padding:18px 24px;border-radius:12px;
    margin-bottom:16px;box-shadow:0 2px 10px rgba(0,0,0,.15);
  }
  .qc-hdr h1 { font-size:22px;margin:0;font-weight:700; }
  .qc-hdr .sub { font-size:11px;opacity:.55;margin-top:2px;
                 text-transform:uppercase;letter-spacing:.4px; }
  .ctx-pill {
    background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.13);
    border-radius:8px;padding:8px 14px;display:inline-block;margin:4px 6px 0 0;
  }
  .ctx-pill .lbl { font-size:10px;opacity:.55;display:block;
                   text-transform:uppercase;letter-spacing:.4px; }
  .ctx-pill .val { font-size:12px;font-weight:600;color:#fff; }
  .ctx-pill.hl  { background:rgba(251,191,36,.15);border-color:rgba(251,191,36,.4); }
  .ctx-pill.hl .val { color:#fbbf24; }
  .badge-auto { background:#3b82f6;color:#fff;font-size:10px;font-weight:700;
                padding:3px 10px;border-radius:10px; }
  .badge-man  { background:#f59e0b;color:#1a2942;font-size:10px;font-weight:700;
                padding:3px 10px;border-radius:10px; }
  .pill-err { background:#fee2e2;color:#dc2626;padding:2px 8px;border-radius:4px;
              font-size:10px;font-weight:700;letter-spacing:.3px; }
  .pill-pep { background:#ffedd5;color:#c2410c;padding:2px 8px;border-radius:4px;
              font-size:10px;font-weight:700;letter-spacing:.3px; }
  .pill-wrn { background:#fef3c7;color:#b45309;padding:2px 8px;border-radius:4px;
              font-size:10px;font-weight:700;letter-spacing:.3px; }
  .pill-apr { background:#ede9fe;color:#6d28d9;padding:2px 8px;border-radius:4px;
              font-size:10px;font-weight:700;letter-spacing:.3px; }
</style>
""", unsafe_allow_html=True)


# ─── SESSION STATE ──────────────────────────────────────────────────
if 'results' not in st.session_state:
    st.session_state.results = None
if 'filter' not in st.session_state:
    st.session_state.filter = 'all'


# ─── MONTH PICKER ───────────────────────────────────────────────────
auto_m, auto_y = auto_detect_month()
today = datetime.now()

options = []
for i in range(-18, 3):
    yy = today.year
    mm = today.month + i
    while mm > 12: mm -= 12; yy += 1
    while mm < 1:  mm += 12; yy -= 1
    options.append((mm - 1, yy))

default_idx = next((i for i, (m, y) in enumerate(options)
                    if m == auto_m and y == auto_y), 18)

# ─── HEADER ─────────────────────────────────────────────────────────
st.markdown(
    '<div class="qc-hdr">'
    '<h1>📋 Invoice QC Tool</h1>'
    '<div class="sub">Redirect Health · Automated Quality Control</div>'
    '</div>', unsafe_allow_html=True,
)

col_l, col_r = st.columns([3, 2])
with col_r:
    sel = st.selectbox(
        "Invoice Month",
        options=options,
        index=default_idx,
        format_func=lambda x: f"{MF[x[0]]} {x[1]}",
    )
sel_m, sel_y = sel
is_auto = (sel_m == auto_m and sel_y == auto_y)
ctx = build_context(sel_m, sel_y)

with col_l:
    badge = ('<span class="badge-auto">AUTO</span>' if is_auto
             else '<span class="badge-man">MANUAL</span>')
    pills = [
        ('Prep Window',     ctx['prepLabel'],      False),
        ('Data Window',     ctx['windowLabel'],    False),
        ('60-Day Free',     ctx['freeLabel'],      False),
        ('Needs Approval',  ctx['approvalLabel'],  not is_auto),
    ]
    pills_html = ''.join(
        f'<div class="ctx-pill{" hl" if hl else ""}">'
        f'<span class="lbl">{lbl}</span><span class="val">{val}</span></div>'
        for lbl, val, hl in pills
    )
    st.markdown(
        f'<div style="background:#1a2942;padding:14px 18px;border-radius:10px;">'
        f'<div style="margin-bottom:6px;">{badge} '
        f'<span style="color:#fff;font-weight:700;font-size:14px;margin-left:10px;">'
        f'{ctx["label"]} Invoice</span></div>'
        f'{pills_html}</div>', unsafe_allow_html=True,
    )

st.write("")


# ─── UPLOAD STAGE ───────────────────────────────────────────────────
if st.session_state.results is None:
    st.subheader(f"📁 Upload Files to QC — {ctx['label']} Invoice")

    col_excel, col_pdfs = st.columns([1, 1.6])
    with col_excel:
        st.markdown("**📊 Merged Excel File**")
        st.caption("EN + SFTP merged data (.xlsx)")
        excel_file = st.file_uploader(
            "Excel file", type=['xlsx', 'xls'],
            label_visibility='collapsed', key='excel_up',
        )
        if excel_file:
            st.success(f"✓ {excel_file.name} ({excel_file.size // 1024} KB)")

    with col_pdfs:
        st.markdown("**📄 Invoice PDFs**")
        st.caption("Add one or many — they accumulate")
        pdf_files = st.file_uploader(
            "PDFs", type=['pdf'], accept_multiple_files=True,
            label_visibility='collapsed', key='pdf_up',
        )
        if pdf_files:
            st.success(f"✓ {len(pdf_files)} PDF{'s' if len(pdf_files) != 1 else ''} ready")
            with st.expander("File list", expanded=False):
                for i, f in enumerate(pdf_files, 1):
                    st.text(f"  {i:>2}. {f.name}  ({f.size // 1024} KB)")

    st.write("")
    can_run = (excel_file is not None) and (pdf_files is not None) and len(pdf_files) > 0
    btn_label = (f"▶ Run QC — {ctx['label']}  ({len(pdf_files)} "
                 f"invoice{'s' if len(pdf_files) != 1 else ''})"
                 if pdf_files else "▶ Run QC Check")

    if st.button(btn_label, disabled=not can_run,
                 use_container_width=True, type='primary'):
        try:
            with st.spinner("Parsing Excel file..."):
                df = pd.read_excel(excel_file)
                rows = df.to_dict('records')

            groups = defaultdict(list)
            for r in rows:
                gid = str(r.get('Identifier', '') or '').strip()
                groups[gid].append(r)

            results = []
            prog = st.progress(0.0, text="Starting...")
            for i, pdf_file in enumerate(pdf_files):
                prog.progress(i / len(pdf_files),
                              text=f"Extracting: {pdf_file.name}")
                try:
                    pdf_data = extract_pdf(pdf_file)
                    gid = pdf_data.get('groupId', '') or ''
                    qc = run_qc(groups.get(gid, []), pdf_data, ctx)
                    qc['_pdf_name'] = pdf_file.name
                    qc['_raw_text'] = pdf_data.get('_raw_text', '')
                    qc['_pdf_data'] = pdf_data
                    results.append(qc)
                except Exception as e:
                    st.warning(f"⚠️ Could not parse {pdf_file.name}: {e}")
                    results.append({
                        'groupId': '', 'groupName': f'❌ Parse error: {pdf_file.name}',
                        'invoiceAmount': 0, 'currentPeriodAmount': 0,
                        'adjCharges': 0, 'adjCredits': 0,
                        'mi': [], 'la': [], 'validations': [],
                        'hasErrors': True, 'hasWarnings': False, 'isClean': False,
                        '_pdf_name': pdf_file.name, '_raw_text': '', '_error': str(e),
                    })
                prog.progress((i + 1) / len(pdf_files),
                              text=f"Done {i + 1} / {len(pdf_files)}")

            prog.empty()
            st.session_state.results = results
            st.rerun()

        except Exception as e:
            st.error(f"⚠️ Error: {e}")


# ─── RESULTS STAGE ──────────────────────────────────────────────────
else:
    R = st.session_state.results

    err_c   = sum(1 for r in R if r['hasErrors'])
    warn_c  = sum(1 for r in R if r['hasWarnings'])
    clean_c = sum(1 for r in R if r['isClean'])
    late_c  = sum(len(r['la']) for r in R)
    total_a = sum((r.get('invoiceAmount') or 0) for r in R)

    # Validation tallies across all groups
    all_validations = [v for r in R for v in r.get('validations', [])]
    v_ok    = sum(1 for v in all_validations if v['status'] == 'ok')
    v_flag  = sum(1 for v in all_validations if v['status'] == 'flag')
    v_wrong = sum(1 for v in all_validations if v['status'] == 'incorrect')

    st.markdown(
        f'<div style="background:#f0f4ff;border:1px solid #c7d2fe;border-radius:10px;'
        f'padding:10px 16px;margin-bottom:14px;font-size:12px;color:#4338ca;'
        f'display:flex;align-items:center;gap:18px;flex-wrap:wrap;">'
        f'<span style="font-weight:700;font-size:13px;color:#1a2942;">'
        f'Results: {ctx["label"]} Invoice</span>'
        f'<span>Data Window: <strong>{ctx["windowLabel"]}</strong></span>'
        f'<span>60-Day Free: <strong>{ctx["freeLabel"]}</strong></span>'
        f'</div>', unsafe_allow_html=True,
    )

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("🔴 Critical Errors", err_c)
    c2.metric("🟡 Warnings",        warn_c)
    c3.metric("⚠️ Late (Approval)", late_c)
    c4.metric("✅ Clean Groups",    clean_c)
    c5.metric("Total Value",        f"${total_a:,.0f}")

    # Adjustment validation summary row
    if all_validations:
        v1, v2, v3, v4 = st.columns(4)
        v1.metric("📊 PDF Adjustments", len(all_validations))
        v2.metric("✅ OK",         v_ok)
        v3.metric("🟡 Flagged",    v_flag)
        v4.metric("🔴 Incorrect",  v_wrong)

    f_col, _, reset_col = st.columns([4, 4, 1.5])
    with f_col:
        filt = st.radio(
            "Filter",
            options=['all', 'errors', 'warnings', 'clean'],
            format_func=lambda x: {
                'all': f'All ({len(R)})',
                'errors': f'Errors ({err_c})',
                'warnings': f'Warnings ({warn_c})',
                'clean': f'Clean ({clean_c})',
            }[x],
            horizontal=True,
            index=['all','errors','warnings','clean'].index(st.session_state.filter),
            label_visibility='collapsed',
        )
        st.session_state.filter = filt
    with reset_col:
        if st.button("↺ Start Over", use_container_width=True):
            st.session_state.results = None
            st.session_state.filter = 'all'
            st.rerun()

    fil = list(R)
    if filt == 'errors':     fil = [r for r in R if r['hasErrors']]
    elif filt == 'warnings': fil = [r for r in R if r['hasWarnings']]
    elif filt == 'clean':    fil = [r for r in R if r['isClean']]
    fil.sort(key=lambda r: 0 if r['hasErrors'] else 1 if r['hasWarnings'] else 2)

    # ── PER-GROUP RENDER ────────────────────────────────────────────
    for g in fil:
        if g['hasErrors']:    tag, dot = "🔴 ERRORS",   "🔴"
        elif g['hasWarnings']: tag, dot = "🟡 WARNINGS", "🟡"
        else:                 tag, dot = "🟢 CLEAN",    "🟢"
        n_iss = len(g['mi']) + len(g['la'])
        n_val_issues = sum(1 for v in g.get('validations', [])
                           if v['status'] in ('flag', 'incorrect'))
        amt = g.get('invoiceAmount') or 0
        title = (f"{dot}  {g.get('groupName') or 'Unknown'}  "
                 f"·  {g.get('groupId') or '—'}  ·  ${amt:,.0f}")
        if n_iss:        title += f"  ·  {n_iss} issue{'s' if n_iss != 1 else ''}"
        if n_val_issues: title += f"  ·  {n_val_issues} adj-flag{'s' if n_val_issues != 1 else ''}"
        title += f"  ·  {tag}"

        with st.expander(title, expanded=g['hasErrors']):
            # Amount breakdown
            a1, a2, a3, a4 = st.columns(4)
            a1.metric("Current Period", f"${(g.get('currentPeriodAmount') or 0):,.0f}")
            a2.metric("Adj Charges",    f"${(g.get('adjCharges') or 0):,.0f}")
            credits = g.get('adjCredits') or 0
            a3.metric("Adj Credits", f"-${abs(credits):,.0f}" if credits else "$0")
            a4.metric("Invoice Total", f"${amt:,.0f}")

            if g.get('_error'):
                st.error(f"Parse error: {g['_error']}")

            # Member issues
            if g['mi']:
                st.markdown("##### Member Issues")
                for m in g['mi']:
                    st.markdown(f"**👤 {m['name']}**")
                    for i in m['iss']:
                        pill_class = {'error': 'pill-err', 'pepm': 'pill-pep',
                                      'warning': 'pill-wrn', 'approval': 'pill-apr'}[i['sev']]
                        label = {'error': 'ERROR', 'pepm': 'PEPM',
                                 'warning': 'WARN', 'approval': 'APPROVAL'}[i['sev']]
                        st.markdown(
                            f'<div style="margin:4px 0 4px 14px;">'
                            f'<span class="{pill_class}">{label}</span> '
                            f'<span style="font-size:13px;color:#374151;">{i["msg"]}</span>'
                            f'</div>', unsafe_allow_html=True,
                        )

            # ── NEW: PDF Adjustment Validation Table ────────────────
            vals = g.get('validations', [])
            if vals:
                st.markdown("##### 📋 PDF Adjustment Validation")

                ok_n    = sum(1 for v in vals if v['status'] == 'ok')
                fl_n    = sum(1 for v in vals if v['status'] == 'flag')
                wrong_n = sum(1 for v in vals if v['status'] == 'incorrect')
                st.caption(f"✅ {ok_n} OK · 🟡 {fl_n} flagged · 🔴 {wrong_n} incorrect")

                v_df = pd.DataFrame([{
                    'Status':   {'ok': '✅ OK', 'flag': '🟡 FLAG',
                                 'incorrect': '🔴 INCORRECT'}[v['status']],
                    'Type':     v['type'],
                    'Member':   v['name'],
                    'Coverage Month': v['month'],
                    'Plan':     v.get('planCode', ''),
                    'Amount':   f"${abs(v['cost']):,.2f}" + (' (cr)' if v['type'] == 'Credit' else ''),
                    'Reason':   v['reason'],
                } for v in vals])

                def _row_color(row):
                    if row['Status'].startswith('🔴'): bg = '#fee2e2'
                    elif row['Status'].startswith('🟡'): bg = '#fef3c7'
                    else: bg = '#dcfce7'
                    return [f'background-color: {bg}'] * len(row)

                st.dataframe(
                    v_df.style.apply(_row_color, axis=1),
                    use_container_width=True, hide_index=True,
                )

            if g['isClean'] and not g['la'] and not vals:
                st.success("✅ All checks passed — no issues found")

            # Late adjustments
            if g['la']:
                st.warning("⚠️ Late Adjustments — beyond 60 days, approval required")
                la_df = pd.DataFrame(g['la'])
                la_df['cost'] = la_df['cost'].apply(lambda x: f"${abs(x):,.2f}")
                st.dataframe(la_df, use_container_width=True, hide_index=True)

            # Debug — raw PDF text
            with st.expander("🔎 Debug — raw extracted text from this PDF"):
                st.text(g.get('_raw_text', '') or '(no text extracted)')

    # ── EXPORTS ─────────────────────────────────────────────────────
    st.divider()
    st.subheader("Export Reports")

    issues_rows = []
    for g in R:
        for m in g['mi']:
            for i in m['iss']:
                issues_rows.append({
                    'Group ID':    g['groupId'],
                    'Group Name':  g['groupName'],
                    'Member':      m['name'],
                    'Severity':    i['sev'].upper(),
                    'Message':     i['msg'],
                    'Invoice Amt': g.get('invoiceAmount') or 0,
                })
        for a in g['la']:
            issues_rows.append({
                'Group ID':    g['groupId'],
                'Group Name':  g['groupName'],
                'Member':      a['name'],
                'Severity':    f"LATE {a['type'].upper()}",
                'Message':     f"{a['month']} · {a['plan']} · ${abs(a['cost']):,.2f}",
                'Invoice Amt': g.get('invoiceAmount') or 0,
            })

    valid_rows = []
    for g in R:
        for v in g.get('validations', []):
            valid_rows.append({
                'Group ID':       g['groupId'],
                'Group Name':     g['groupName'],
                'Member':         v['name'],
                'Adjustment Type': v['type'],
                'Coverage Month': v['month'],
                'Plan':           v.get('planCode', ''),
                'Amount':         v['cost'],
                'Status':         v['status'].upper(),
                'Reason':         v['reason'],
            })

    e_col1, e_col2 = st.columns(2)
    with e_col1:
        if issues_rows:
            csv_bytes = pd.DataFrame(issues_rows).to_csv(index=False).encode('utf-8')
            st.download_button(
                "⬇️ Issues report (CSV)",
                data=csv_bytes,
                file_name=f"qc_issues_{ctx['label'].replace(' ', '_')}.csv",
                mime='text/csv', use_container_width=True,
            )
        else:
            st.info("No Excel-driven issues to export 🎉")

    with e_col2:
        if valid_rows:
            csv_bytes = pd.DataFrame(valid_rows).to_csv(index=False).encode('utf-8')
            st.download_button(
                "⬇️ Adjustment validation report (CSV)",
                data=csv_bytes,
                file_name=f"adjustment_validation_{ctx['label'].replace(' ', '_')}.csv",
                mime='text/csv', use_container_width=True,
            )
        else:
            st.info("No PDF adjustments to validate")
