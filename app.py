import streamlit as st
import pandas as pd
import xlsxwriter
import io
import os

# ── page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Investor Statement Extractor",
    page_icon="📊",
    layout="centered"
)

# ── styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

  html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

  .main { background-color: #0f2340; }
  .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 780px; }

  .hero {
    background: linear-gradient(135deg, #1a3355 0%, #0f2340 100%);
    border-radius: 16px;
    padding: 2rem 2.5rem;
    margin-bottom: 2rem;
    border: 1px solid rgba(46,125,255,0.2);
  }
  .hero-eyebrow {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    letter-spacing: 0.18em;
    color: #2e7dff;
    text-transform: uppercase;
    margin-bottom: 8px;
  }
  .hero h1 {
    font-size: 2rem;
    font-weight: 600;
    color: #ffffff;
    margin: 0 0 8px 0;
    letter-spacing: -0.02em;
  }
  .hero h1 span { color: #2e7dff; }
  .hero p { color: #8a9bb5; font-size: 14px; margin: 0; line-height: 1.6; }

  .fund-tag {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 5px;
    font-family: 'DM Mono', monospace;
    font-size: 12px;
    font-weight: 500;
    margin: 2px 4px 2px 0;
  }
  .tag-cpdf   { background: rgba(46,125,255,0.15); color: #2e7dff; border: 1px solid rgba(46,125,255,0.3); }
  .tag-dlot   { background: rgba(245,166,35,0.15); color: #f5a623; border: 1px solid rgba(245,166,35,0.3); }
  .tag-cdlot2 { background: rgba(29,184,122,0.15); color: #1db87a; border: 1px solid rgba(29,184,122,0.3); }
  .tag-other  { background: rgba(138,155,181,0.15); color: #8a9bb5; border: 1px solid rgba(138,155,181,0.3); }

  .stat-card {
    background: #1a3355;
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 12px;
    padding: 1rem 1.2rem;
    text-align: center;
  }
  .stat-label {
    font-family: 'DM Mono', monospace;
    font-size: 10px;
    letter-spacing: 0.12em;
    color: #8a9bb5;
    text-transform: uppercase;
    margin-bottom: 4px;
  }
  .stat-val { font-size: 28px; font-weight: 600; letter-spacing: -0.02em; }
  .stat-blue  { color: #2e7dff; }
  .stat-green { color: #1db87a; }
  .stat-white { color: #ffffff; }
  .stat-amber { color: #f5a623; }

  .file-row {
    display: flex;
    align-items: center;
    gap: 10px;
    background: #1a3355;
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 8px;
    padding: 10px 14px;
    margin-bottom: 6px;
    font-size: 13px;
    color: #e8edf5;
  }
  .file-name { flex: 1; font-family: 'DM Mono', monospace; font-size: 12px; }

  .success-box {
    background: rgba(29,184,122,0.1);
    border: 1px solid rgba(29,184,122,0.3);
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    margin: 1rem 0;
    color: #1db87a;
    font-size: 14px;
  }
  .warning-box {
    background: rgba(245,166,35,0.1);
    border: 1px solid rgba(245,166,35,0.3);
    border-radius: 10px;
    padding: 0.8rem 1.2rem;
    color: #f5a623;
    font-size: 13px;
    margin: 6px 0;
  }
  .section-title {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    letter-spacing: 0.14em;
    color: #8a9bb5;
    text-transform: uppercase;
    margin: 1.5rem 0 0.6rem 0;
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .section-title::after {
    content: '';
    flex: 1;
    height: 1px;
    background: rgba(255,255,255,0.08);
  }

  /* override streamlit upload button */
  [data-testid="stFileUploader"] {
    background: #1a3355;
    border: 1.5px dashed rgba(46,125,255,0.35);
    border-radius: 12px;
    padding: 0.5rem;
  }
  [data-testid="stFileUploader"]:hover {
    border-color: #2e7dff;
  }

  /* download button */
  [data-testid="stDownloadButton"] > button {
    background: #1db87a !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.6rem 2rem !important;
    font-size: 15px !important;
    font-weight: 600 !important;
    width: 100%;
  }
  [data-testid="stDownloadButton"] > button:hover {
    opacity: 0.88 !important;
  }
</style>
""", unsafe_allow_html=True)

# ── constants ─────────────────────────────────────────────────────────────────
FUND_CODES = ['CDLOT2', 'CPDF', 'DLOT', 'CDLOT']

SUFFIXES = [
    'Superannuation Fund', 'Superannunation Fund', 'Supperannuation Fund',
    'Superannuation', 'Superannunation', 'Supperannuation',
    'Trust Fund', 'Family Trust', 'Trust', 'Fund'
]

MASTER = {
    'John Barnett Nominees Pty Ltd ATF JGB Superannunation Fund':
        'John Barnett Nominees Pty Ltd ATF JGB Superannuation Fund',
    'Hunter Brunelle Nominees Pty Ltd RS Hunter Supperannuation Fund':
        'Hunter Brunelle Nominees Pty Ltd RS Hunter Superannuation Fund',
    'IT Contracting Services Pty Ltd ATF Zabow Supperannuation Fund':
        'IT Contracting Services Pty Ltd ATF Zabow Superannuation Fund',
    'Adorben Pty Limited ATF Adorben Pty Limited Supperannuation Fund':
        'Adorben Pty Limited ATF Adorben Pty Limited',
    'Kaufline Superannuation Pty Ltd ATF Kaufline Family':
        'Kaufline Superannuation Pty Ltd ATF Kaufline Family Super Fund',
    'BDH Superannuation Fund Pty Ltd ATF BDH Supperannuation Fund':
        'BDH Superannuation Fund Pty Ltd ATF BDH Superannuation Fund',
    'D2 Enterprises Pty Ltd ATF The Muirhead Supperannuation Fund':
        'D2 Enterprises Pty Ltd ATF The Muirhead Superannuation Fund',
    'Graymere Pty Limited ATF The Graymere Superannuation Fund':
        'Graymere Pty Limited ATF The Graymere Superannuation',
    'Loreak Mendian Pty Ltd ATF Telleria Family Trust': 'Richard Telleria',
    'Sesame Bagel Pty Ltd ATF Sesame Bagel Trust': 'Sesame Bagel Trust',
    'Gandalf Investments Pty Ltd ATF Elliot Rubinstein Supperannuation':
        'Gandalf Investments Pty Ltd ATF Elliot Rubinstein',
    'Supermann Pty Ltd ATF Cartisano Super Fund':
        'Supermann Pty Limited ATF Cartisano Superannuation Fund',
    'Abata Pty Ltd ATF Williams Family Trust': 'Abata Pty Ltd',
    'Constel Investments Pty Ltd ATF Pavlakos Family Super Fund':
        'Constel Investments Pty Ltd ATF Pavlakos Family Superannuation F',
}

# ── helpers ───────────────────────────────────────────────────────────────────
def detect_entity(filename):
    name = filename.upper()
    for code in FUND_CODES:
        if code in name:
            return code
    return 'UNKNOWN'

def cv(v):
    if v is None:
        return ''
    s = str(v).strip()
    return '' if s == 'nan' else s

def extract_file(uploaded_file):
    entity = detect_entity(uploaded_file.name)
    df     = pd.read_excel(uploaded_file, header=None)
    rows   = df.values.tolist()
    results = []

    for i, row in enumerate(rows):
        if cv(row[1]) != 'CERTIFICATE HOLDER':
            continue

        balance = 0.0
        for col in [23, 18, 16]:
            raw = row[col] if col < len(row) else None
            if raw is not None and str(raw).strip() not in ('', 'nan'):
                try:
                    balance = float(raw)
                    break
                except (ValueError, TypeError):
                    pass

        name_row   = rows[i + 2] if i + 2 < len(rows) else []
        suffix_row = rows[i + 3] if i + 3 < len(rows) else []
        name = cv(name_row[1]) if len(name_row) > 1 else ''
        if not name:
            continue

        suffix = cv(suffix_row[1]) if len(suffix_row) > 1 else ''
        if suffix in SUFFIXES:
            name = name + ' ' + suffix
        name     = name.strip()
        investor = MASTER.get(name, name)
        results.append({'entity': entity, 'investor': investor, 'balance': balance})

    return entity, results

def build_excel(all_results):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Client Statement')

    hdr = wb.add_format({
        'bold': True, 'font_name': 'Arial', 'font_size': 10,
        'bg_color': '#1F4E79', 'font_color': '#FFFFFF',
        'align': 'center', 'valign': 'vcenter'
    })
    body  = wb.add_format({'font_name': 'Arial', 'font_size': 10})
    money = wb.add_format({
        'font_name': 'Arial', 'font_size': 10,
        'num_format': '#,##0.00;-#,##0.00;"-"',
        'align': 'right'
    })

    for c, h in enumerate(['Entity', 'Investor', 'Entity | Investor', 'Balance']):
        ws.write(0, c, h, hdr)

    for r, rec in enumerate(all_results, 1):
        ws.write(r, 0, rec['entity'], body)
        ws.write(r, 1, rec['investor'], body)
        ws.write(r, 2, rec['entity'] + ' | ' + rec['investor'], body)
        ws.write(r, 3, rec['balance'], money)

    ws.set_column('A:A', 10)
    ws.set_column('B:B', 58)
    ws.set_column('C:C', 70)
    ws.set_column('D:D', 18)
    ws.set_row(0, 18)
    wb.close()
    output.seek(0)
    return output

def tag_html(entity):
    cls = {
        'CPDF': 'tag-cpdf', 'DLOT': 'tag-dlot',
        'CDLOT2': 'tag-cdlot2'
    }.get(entity, 'tag-other')
    return f'<span class="fund-tag {cls}">{entity}</span>'

# ── app ───────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-eyebrow">Capspace · Investor Tools</div>
  <h1>Statement <span>Extractor</span></h1>
  <p>Upload one or more CPDF, DLOT, or CDLOT2 investor statement files.<br>
     All investors are combined into a single downloadable Excel sheet.</p>
</div>
""", unsafe_allow_html=True)

# ── upload ────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Upload files</div>', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    label="Drop your Excel files here or click to browse",
    type=['xlsx', 'xls'],
    accept_multiple_files=True,
    help="Select the raw monthly Investor_Statements_*.xlsx files — not the extracted output."
)

if uploaded_files:
    # show file list with detected fund tags
    st.markdown('<div class="section-title">Detected files</div>', unsafe_allow_html=True)
    all_unknown = False
    for f in uploaded_files:
        entity = detect_entity(f.name)
        if entity == 'UNKNOWN':
            all_unknown = True
        st.markdown(
            f'<div class="file-row">{tag_html(entity)}'
            f'<span class="file-name">{f.name}</span>'
            f'<span style="color:#8a9bb5;font-size:12px">'
            f'{round(f.size/1024)}KB</span></div>',
            unsafe_allow_html=True
        )

    if all_unknown:
        st.markdown(
            '<div class="warning-box">⚠  Could not detect fund from filename. '
            'Make sure the filename contains CPDF, DLOT, or CDLOT2.</div>',
            unsafe_allow_html=True
        )

    # ── extract ───────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">Extract</div>', unsafe_allow_html=True)

    if st.button("   ⚡  Extract all files   ", use_container_width=True):
        all_results  = []
        file_summary = []
        errors       = []

        progress = st.progress(0, text="Starting…")

        for idx, f in enumerate(uploaded_files):
            progress.progress(
                int((idx / len(uploaded_files)) * 90),
                text=f"Reading {f.name}…"
            )
            try:
                entity, results = extract_file(f)
                all_results.extend(results)
                file_summary.append((entity, f.name, len(results)))
            except Exception as e:
                errors.append(f"{f.name}: {str(e)}")

        progress.progress(100, text="Building output…")

        if errors:
            for err in errors:
                st.error(err)

        if not all_results:
            st.error(
                "No investor data found. Make sure you uploaded the raw "
                "statement files, not the extracted output."
            )
            progress.empty()
        else:
            excel_data = build_excel(all_results)
            progress.empty()

            # summary stats
            non_zero = sum(1 for r in all_results if r['balance'] != 0)
            total    = sum(r['balance'] for r in all_results)

            st.markdown('<div class="section-title">Results</div>', unsafe_allow_html=True)

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.markdown(
                    f'<div class="stat-card"><div class="stat-label">Files</div>'
                    f'<div class="stat-val stat-amber">{len(uploaded_files)}</div></div>',
                    unsafe_allow_html=True)
            with c2:
                st.markdown(
                    f'<div class="stat-card"><div class="stat-label">Investors</div>'
                    f'<div class="stat-val stat-blue">{len(all_results)}</div></div>',
                    unsafe_allow_html=True)
            with c3:
                st.markdown(
                    f'<div class="stat-card"><div class="stat-label">Non-zero</div>'
                    f'<div class="stat-val stat-white">{non_zero}</div></div>',
                    unsafe_allow_html=True)
            with c4:
                st.markdown(
                    f'<div class="stat-card"><div class="stat-label">Total Balance</div>'
                    f'<div class="stat-val stat-green">${total:,.0f}</div></div>',
                    unsafe_allow_html=True)

            # per-file breakdown
            st.markdown('<br>', unsafe_allow_html=True)
            for entity, fname, count in file_summary:
                st.markdown(
                    f'<div class="file-row">{tag_html(entity)}'
                    f'<span class="file-name">{fname}</span>'
                    f'<span style="color:#1db87a;font-family:\'DM Mono\',monospace;font-size:12px">'
                    f'{count} investors</span></div>',
                    unsafe_allow_html=True
                )

            st.markdown(
                '<div class="success-box">✓  Extraction complete — click below to download.</div>',
                unsafe_allow_html=True
            )

            # download button
            st.markdown('<div style="margin-top:1rem">', unsafe_allow_html=True)
            st.download_button(
                label="⬇  Download Combined_Statement_Extracted.xlsx",
                data=excel_data,
                file_name="Combined_Statement_Extracted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.markdown('</div>', unsafe_allow_html=True)

            # preview table
            st.markdown('<div class="section-title">Preview (first 20 rows)</div>',
                        unsafe_allow_html=True)
            preview = pd.DataFrame(all_results[:20])[['entity', 'investor', 'balance']]
            preview.columns = ['Entity', 'Investor', 'Balance']
            preview['Balance'] = preview['Balance'].apply(
                lambda x: f"${x:,.2f}" if x != 0 else "-"
            )
            st.dataframe(preview, use_container_width=True, hide_index=True)

else:
    st.markdown("""
    <div style="text-align:center;padding:2rem;color:#8a9bb5;font-size:14px">
        <div style="font-size:36px;margin-bottom:12px">📂</div>
        Upload your investor statement Excel files above to get started.<br>
        <span style="font-family:'DM Mono',monospace;font-size:12px;color:#2e7dff">
        Supports: CPDF · DLOT · CDLOT2</span>
    </div>
    """, unsafe_allow_html=True)

# ── footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="margin-top:3rem;padding-top:1.5rem;
     border-top:1px solid rgba(255,255,255,0.08);
     text-align:center;font-size:12px;color:#8a9bb5;
     font-family:'DM Mono',monospace;">
  Capspace · Investor Statement Extractor
</div>
""", unsafe_allow_html=True)
