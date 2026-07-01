"""
app.py — Serials Holdings Cataloging Tool
Hood College Library

To run:  streamlit run app.py
"""

import streamlit as st
from helpers import (
    extract_issn, get_journal_info, get_all_works, works_to_issues,
    sort_issues, parse_manual_text, merge_issues,
    issues_to_dataframe, dataframe_to_excel, generate_claude_prompt,
    filter_by_holdings_range, get_volume_year_map,
)
from marc_helpers import (
    detect_format, parse_phase1_upload, parse_alma_export,
    normalize_session_issues, analyze_pattern, generate_marc_fields, generate_report,
)

# ── PAGE CONFIG ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title='Serials Cataloging Tool — Hood College Library',
    page_icon='📚',
    layout='wide'
)

# ── CUSTOM CSS ───────────────────────────────────────────────────────────────

st.markdown("""
<style>
/* Tighten top padding */
.block-container { padding-top: 1.5rem; }

/* Home screen card styling */
.tool-card {
    background: #F2EDE5;
    border: 1px solid #C9B99A;
    border-radius: 8px;
    padding: 1.75rem 1.75rem 1.25rem 1.75rem;
    margin-bottom: 0.6rem;
    min-height: 180px;
}
.tool-card-icon { font-size: 1.6rem; margin-bottom: 0.5rem; }
.tool-card-title {
    font-weight: 700;
    font-size: 1.1rem;
    color: #2C2C2C;
    margin: 0 0 0.6rem 0;
}
.tool-card-body {
    color: #5A4F4A;
    line-height: 1.65;
    font-size: 0.95rem;
    margin: 0;
}
.tool-card-phase {
    margin-top: 1rem;
    font-size: 0.8rem;
    font-weight: 700;
    color: #7A2035;
    letter-spacing: 0.05em;
    text-transform: uppercase;
}
/* Home header */
.home-header {
    text-align: center;
    padding: 2rem 0 2rem 0;
    border-bottom: 1px solid #C9B99A;
    margin-bottom: 2rem;
}
.home-title {
    font-size: 1.9rem;
    font-weight: 700;
    color: #7A2035;
    margin: 0 0 0.3rem 0;
}
.home-subtitle {
    font-size: 1rem;
    color: #5A4F4A;
    margin: 0;
}
/* Phase badge shown in wizard steps */
.phase-header {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    margin-bottom: 0.5rem;
}
.phase-tag {
    display: inline-block;
    background: #7A2035;
    color: white;
    padding: 0.2rem 0.65rem;
    border-radius: 4px;
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 0.04em;
    text-transform: uppercase;
}
</style>
""", unsafe_allow_html=True)

# ── SESSION STATE ────────────────────────────────────────────────────────────

DEFAULTS = {
    'step': 0,
    # Phase 1
    'issn': None,
    'journal_title': '',
    'crossref_issues': [],
    'manual_text': '',
    'all_issues': [],
    'mms_id': '',
    'holding_id': '',
    'df': None,
    'crossref_done': False,
    'holdings_start_vol': None,
    'holdings_end_vol': None,
    'special_entries': [],
    # Phase 2
    'p2_issues': [],
    'p2_journal_title': '',
    'p2_issn': '',
    'p2_analysis': None,
    'p2_marc_sections': [],
    'p2_report': '',
}

for key, val in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = val


def go(step: int):
    st.session_state.step = step


def show_home_link():
    """Show a subtle Home link above the progress bar."""
    if st.button('⌂  Home', key='home_link'):
        go(0)
        st.rerun()


def show_phase1_progress():
    STEPS = ['ISSN', 'Confirm', 'Discover', 'Alma IDs', 'Review', 'Download']
    step = st.session_state.step
    show_home_link()
    st.markdown('<span class="phase-tag">Phase 1 — Item Records</span>', unsafe_allow_html=True)
    st.progress((step - 1) / (len(STEPS) - 1))
    cols = st.columns(len(STEPS))
    for i, label in enumerate(STEPS):
        if i + 1 == step:
            cols[i].markdown(
                f'<span style="color:#7A2035;font-weight:700">{i + 1}. {label}</span>',
                unsafe_allow_html=True
            )
        else:
            cols[i].markdown(f'{i + 1}. {label}')
    st.divider()


def show_phase2_progress():
    STEPS = ['Load', 'Analysis', 'MARC Fields', 'Report']
    step = st.session_state.step - 6   # step 7 → position 1, etc.
    show_home_link()
    st.markdown('<span class="phase-tag">Phase 2 — MARC Pattern Generator</span>', unsafe_allow_html=True)
    st.progress((step - 1) / (len(STEPS) - 1))
    cols = st.columns(len(STEPS))
    for i, label in enumerate(STEPS):
        if i + 1 == step:
            cols[i].markdown(
                f'<span style="color:#7A2035;font-weight:700">{i + 1}. {label}</span>',
                unsafe_allow_html=True
            )
        else:
            cols[i].markdown(f'{i + 1}. {label}')
    st.divider()


# ═══════════════════════════════════════════════════════════════════════════
# STEP 0 — HOME SCREEN
# ═══════════════════════════════════════════════════════════════════════════

if st.session_state.step == 0:
    st.markdown("""
    <div class="home-header">
        <p class="home-title">Serials Holdings Cataloging Tool</p>
        <p class="home-subtitle">Hood College Library</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('#### What would you like to do?')
    st.write('')

    col1, col2 = st.columns(2, gap='large')

    with col1:
        st.markdown("""
        <div class="tool-card">
            <div class="tool-card-icon">📋</div>
            <p class="tool-card-title">Build an items list</p>
            <p class="tool-card-body">
                Start from an ISSN. The tool looks up journal issues via CrossRef,
                lets you fill in any missing records, review and edit everything,
                and exports a spreadsheet ready to import into Alma.
            </p>
            <p class="tool-card-phase">Phase 1 — Item Records Generator</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button('Start with Phase 1 →', type='primary',
                     use_container_width=True, key='btn_p1'):
            go(1)
            st.rerun()

    with col2:
        st.markdown("""
        <div class="tool-card">
            <div class="tool-card-icon">📊</div>
            <p class="tool-card-title">Analyze holdings pattern</p>
            <p class="tool-card-body">
                Already have an approved items list? Upload it and the tool
                detects the publication pattern and generates MARC 21 holdings
                fields (853 / 863 / 866) ready to paste into the Alma
                holdings record.
            </p>
            <p class="tool-card-phase">Phase 2 — MARC Pattern Generator</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button('Start with Phase 2 →', type='primary',
                     use_container_width=True, key='btn_p2'):
            go(7)
            st.rerun()

    st.write('')
    st.caption(
        'Phase 1 must be completed before Phase 2 if you do not already have an items list. '
        'Both phases can be run independently at any time.'
    )


# ═══════════════════════════════════════════════════════════════════════════
# STEP 1 — ISSN / URL INPUT
# ═══════════════════════════════════════════════════════════════════════════

elif st.session_state.step == 1:
    show_phase1_progress()
    st.header('Step 1 — Enter ISSN or Journal URL')
    st.write(
        'Type the ISSN directly (e.g. **0046-4813**), or paste any URL for the journal — '
        'the tool will find the ISSN automatically.'
    )
    st.caption(
        '💡 Journals have two ISSNs: one for the print edition and one for the electronic edition. '
        'CrossRef stores records under the electronic ISSN. If the lookup returns no results, '
        'try the other ISSN — it may be printed on the journal cover or masthead.'
    )

    user_input = st.text_input(
        'ISSN or URL',
        placeholder='0046-4813   or   https://www.tandfonline.com/loi/...',
        label_visibility='collapsed'
    )

    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button('Continue →', type='primary'):
            if not user_input.strip():
                st.error('Please enter an ISSN or a URL.')
            else:
                found = extract_issn(user_input.strip())
                if found:
                    st.session_state.issn = found
                    st.session_state.crossref_done = False
                    go(2)
                    st.rerun()
                else:
                    st.error(
                        'No ISSN found. An ISSN looks like **0046-4813** '
                        '(four digits, a dash, four more digits). '
                        'Please check and try again, or type just the ISSN.'
                    )
    with col2:
        if st.button('← Back to Home'):
            go(0)
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# STEP 2 — CONFIRM ISSN
# ═══════════════════════════════════════════════════════════════════════════

elif st.session_state.step == 2:
    show_phase1_progress()
    st.header('Step 2 — Confirm ISSN')
    st.info(f'The ISSN found is: **{st.session_state.issn}**')
    st.write('Is this the correct ISSN for the journal you want to catalog?')

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button('Yes, correct — continue →', type='primary'):
            go(3)
            st.rerun()
    with col2:
        if st.button('No — go back and correct it'):
            go(1)
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# STEP 3 — CROSSREF DISCOVERY + HOLDINGS RANGE + GAP HANDLING
# ═══════════════════════════════════════════════════════════════════════════

elif st.session_state.step == 3:
    show_phase1_progress()
    st.header('Step 3 — Discover Journal Issues')
    issn = st.session_state.issn

    if not st.session_state.crossref_done:
        with st.spinner(f'Looking up ISSN {issn} in CrossRef — this may take a moment…'):
            info = get_journal_info(issn)
            if info:
                st.session_state.journal_title = info.get('title', '')
            works = get_all_works(issn)
            issues = sort_issues(works_to_issues(works))
            st.session_state.crossref_issues = issues
            st.session_state.crossref_done = True

    issues = st.session_state.crossref_issues
    title = st.session_state.journal_title

    if title:
        st.subheader(f'"{title}"')

    if issues:
        years = [i['year'] for i in issues if i['year']]
        first_yr = min(years) if years else '?'
        last_yr = max(years) if years else '?'
        st.success(
            f'CrossRef found **{len(issues)} issues** for this journal, '
            f'ranging from **{first_yr}** to **{last_yr}**.'
        )
        with st.expander('Preview first 10 issues found'):
            st.table([
                {'Volume': i['volume'], 'Issue': i['issue'],
                 'Year': i['year'], 'Month/Season': i['month']}
                for i in issues[:10]
            ])
    else:
        st.warning(
            '**No records found in CrossRef for this ISSN.** '
            'This is common for older or print-only journals — CrossRef mainly covers '
            'electronic journals with registered DOIs. '
            'You can add all records manually below.'
        )

    st.divider()

    start_vol_sel = None
    end_vol_sel = None

    if issues:
        vol_year_map = get_volume_year_map(issues)
        vols_sorted = sorted(vol_year_map.keys())

        if len(vols_sorted) > 1:
            st.subheader('Hood College holdings range')
            st.write(
                "CrossRef found the full published run. Select the **first** and **last** volume "
                "that Hood College physically owns — records outside this range will be excluded."
            )

            saved_start = st.session_state.holdings_start_vol
            saved_end = st.session_state.holdings_end_vol
            start_default = vols_sorted.index(saved_start) if saved_start in vols_sorted else 0
            end_default = vols_sorted.index(saved_end) if saved_end in vols_sorted else len(vols_sorted) - 1

            col1, col2 = st.columns(2)
            with col1:
                start_vol_sel = st.selectbox(
                    'First volume Hood College owns',
                    options=vols_sorted,
                    format_func=lambda v: f'v.{v}  ({vol_year_map[v]})' if vol_year_map.get(v) else f'v.{v}',
                    index=start_default,
                )
            with col2:
                end_vol_sel = st.selectbox(
                    'Last volume Hood College owns',
                    options=vols_sorted,
                    format_func=lambda v: f'v.{v}  ({vol_year_map[v]})' if vol_year_map.get(v) else f'v.{v}',
                    index=end_default,
                )

            if start_vol_sel is not None and end_vol_sel is not None:
                kept = filter_by_holdings_range(issues, start_vol_sel, end_vol_sel)
                excluded_count = len(issues) - len(kept)
                if excluded_count > 0:
                    st.caption(
                        f'{len(kept)} records within this range · '
                        f'{excluded_count} records outside the range will be excluded.'
                    )
                else:
                    st.caption(f'All {len(issues)} CrossRef records are within this range.')

            st.caption(
                '📝 *Note: this assumes Hood\'s run of the journal has a definite end. '
                'If Hood is still actively receiving this journal, check with your supervisor '
                'about how to handle the open end before importing.*'
            )

    st.divider()

    st.subheader('Add records that CrossRef is missing')
    st.write(
        'If early volumes are missing, or this journal has no electronic version, '
        'you can extract the data from physical journals using Claude.ai and paste it here. '
        'Accepted formats: single month (Apr), combined months (Jan/Feb, Jan-Mar), '
        'or seasons (Spring, Fall/Winter).'
    )

    with st.expander('📋 Get the Claude extraction prompt (click to expand)'):
        prompt_text = generate_claude_prompt(issn, title)
        st.write(
            '**How to use this:**  \n'
            '1. Copy the prompt below  \n'
            '2. Open claude.ai in your browser  \n'
            '3. Upload your photo or scan of the journal  \n'
            '4. Paste the prompt and send  \n'
            '5. Copy Claude\'s response and paste it in the text box below'
        )
        st.code(prompt_text, language=None)

    manual_text = st.text_area(
        'Paste extracted records here — one issue per line  '
        '(e.g.  v.1 n.1 Apr 1969  or  v.2 n.1 Spring 1970  or  v.3 n.1-2 Jan/Feb 1971)',
        value=st.session_state.manual_text,
        height=180,
        placeholder=(
            'v.1 n.1 Apr 1969\n'
            'v.1 n.2 Jan/Feb 1970\n'
            'v.2 n.1 Spring 1971\n'
            'v.2 n.3-4 1972'
        ),
    )
    st.session_state.manual_text = manual_text

    st.divider()

    st.subheader('Add index or supplement entries')
    st.write(
        'Use this for items that are not regular numbered issues: cumulative indexes, '
        'supplements, parts, or other special items. These are rarely in CrossRef.'
    )

    with st.expander('➕ Add an index, supplement, or part entry'):
        with st.form('special_entry_form', clear_on_submit=True):
            col0, col1 = st.columns(2)
            with col0:
                se_type = st.selectbox(
                    'Type',
                    ['Index', 'Supplement', 'Part', 'Special'],
                    help='Index = cumulative index volume. Supplement/Part = accompanying issue.'
                )
                se_vol = st.text_input(
                    'Volume (or volume range for Index)',
                    placeholder='e.g.  1-10  for an index, or  5  for a supplement to v.5',
                )
            with col1:
                se_desc = st.text_input(
                    'Descriptor — leave blank for Index (auto-filled)',
                    placeholder='e.g.  Suppl. 1   or   Part 2',
                    help='For Index this is always "Index" and can be left blank. '
                         'For Supplement/Part, enter the label as it appears on the item.'
                )
                se_year = st.number_input(
                    'Year published', min_value=1800, max_value=2100, value=2000, step=1
                )
                se_month = st.text_input(
                    'Month/Season (optional)',
                    placeholder='Spring, Jan/Feb, Oct — leave blank if unknown',
                )

            submitted = st.form_submit_button('Add this entry')

        if submitted:
            if not se_vol.strip():
                st.error('Volume (or volume range) is required.')
            else:
                issue_val = se_desc.strip() if se_desc.strip() else se_type
                new_entry = {
                    'volume': se_vol.strip(),
                    'issue': issue_val,
                    'year': int(se_year) if se_year else None,
                    'month': se_month.strip(),
                    'is_duplicate': False,
                    'source': 'Manual',
                    'missing': False,
                    'item_type': se_type.lower(),
                }
                st.session_state.special_entries.append(new_entry)
                st.rerun()

        if st.session_state.special_entries:
            st.write('**Entries added so far:**')
            for i, entry in enumerate(st.session_state.special_entries):
                date_str = f"{entry.get('month', '')} {entry.get('year', '')}".strip()
                label = (
                    f"**{entry['item_type'].title()}** — "
                    f"v.{entry['volume']}, {entry['issue']}"
                    + (f' ({date_str})' if date_str else '')
                )
                col_a, col_b = st.columns([11, 1])
                col_a.markdown(label)
                if col_b.button('✕', key=f'del_se_{i}', help='Remove this entry'):
                    st.session_state.special_entries.pop(i)
                    st.rerun()

    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button('Continue with these records →', type='primary'):
            manual = parse_manual_text(manual_text) if manual_text.strip() else []

            if issues and start_vol_sel is not None and end_vol_sel is not None:
                if start_vol_sel <= end_vol_sel:
                    filtered_crossref = filter_by_holdings_range(issues, start_vol_sel, end_vol_sel)
                    st.session_state.holdings_start_vol = start_vol_sel
                    st.session_state.holdings_end_vol = end_vol_sel
                else:
                    st.error('First volume must be less than or equal to the last volume.')
                    st.stop()
            else:
                filtered_crossref = issues

            all_issues = sort_issues(merge_issues(filtered_crossref, manual))
            all_issues = all_issues + list(st.session_state.special_entries)

            if not all_issues:
                st.error(
                    'No records to work with yet. '
                    'Either CrossRef found nothing and no manual records were entered. '
                    'Please add at least some records before continuing.'
                )
            else:
                st.session_state.all_issues = all_issues
                go(4)
                st.rerun()
    with col2:
        if st.button('← Back'):
            go(2)
            st.rerun()

    if issues and manual_text.strip():
        manual_preview = parse_manual_text(manual_text)
        if manual_preview:
            active_crossref = (
                filter_by_holdings_range(issues, start_vol_sel, end_vol_sel)
                if start_vol_sel is not None and end_vol_sel is not None
                else issues
            )
            st.info(
                f'You have **{len(active_crossref)} CrossRef records** and '
                f'**{len(manual_preview)} manual records** — '
                f'they will be merged. Manual records take priority if there is overlap.'
            )


# ═══════════════════════════════════════════════════════════════════════════
# STEP 4 — ALMA CONSTANT FIELDS
# ═══════════════════════════════════════════════════════════════════════════

elif st.session_state.step == 4:
    show_phase1_progress()
    st.header('Step 4 — Enter Alma Record IDs')
    st.write(
        'These two values are the same for every row in the spreadsheet. '
        'Enter them once here. See the **Guide for Alma automation** document '
        'for instructions on finding them in Alma.'
    )

    mms = st.text_input(
        'mms_id',
        value=st.session_state.mms_id,
        placeholder='e.g. 9916229323407581',
        help='The bibliographic record ID. Format must stay as text — do not add spaces.'
    )
    hid = st.text_input(
        'holding_id',
        value=st.session_state.holding_id,
        placeholder='e.g. 22211851990007581',
        help='The holdings record ID. Format must stay as text.'
    )

    st.info(
        '**Tip:** These will be stored as TEXT in the Excel file (not as numbers), '
        'which is what Alma requires. You may see a small warning triangle in Excel — '
        'this is expected and safe to ignore.'
    )

    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button('Continue →', type='primary'):
            if not mms.strip() or not hid.strip():
                st.error('Both mms_id and holding_id are required before continuing.')
            else:
                st.session_state.mms_id = mms.strip()
                st.session_state.holding_id = hid.strip()
                st.session_state.df = issues_to_dataframe(st.session_state.all_issues)
                go(5)
                st.rerun()
    with col2:
        if st.button('← Back'):
            go(3)
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# STEP 5 — REVIEW & EDIT TABLE
# ═══════════════════════════════════════════════════════════════════════════

elif st.session_state.step == 5:
    show_phase1_progress()
    st.header('Step 5 — Review & Edit Item Records')

    df = st.session_state.df
    manual_count = int((df['Source'] == 'Manual').sum()) if 'Source' in df.columns else 0
    crossref_count = int((df['Source'] == 'CrossRef').sum()) if 'Source' in df.columns else 0
    missing_count = int(df['Missing / not owned'].sum()) if 'Missing / not owned' in df.columns else 0
    special_count = int((df['Type'] != 'issue').sum()) if 'Type' in df.columns else 0

    st.write(
        f'**{len(df)} records** — {crossref_count} from CrossRef, '
        f'{manual_count} added manually'
        + (f', {special_count} special (index/supplement)' if special_count else '')
        + '.  \n'
        'Edit any cell directly. '
        'Check **Missing / not owned** for issues Hood College does not hold — '
        'those rows will be excluded from the Excel export. '
        'Check **Has duplicate copy** for issues where the library owns two physical copies.'
    )

    if missing_count > 0:
        st.info(f'{missing_count} rows currently marked as missing / not owned and will be excluded from export.')

    if manual_count > 0:
        st.warning(
            f'{manual_count} records were entered manually. '
            'Please verify these carefully against the physical journal.'
        )

    edited_df = st.data_editor(
        df,
        use_container_width=True,
        num_rows='dynamic',
        column_config={
            'Volume': st.column_config.TextColumn('Volume'),
            'Issue': st.column_config.TextColumn('Issue'),
            'Year': st.column_config.NumberColumn('Year', format='%d', min_value=1800, max_value=2100),
            'Month/Season': st.column_config.TextColumn(
                'Month/Season',
                help='3-letter month (Apr), combined months (Jan/Feb, Jan-Mar), or season (Spring, Fall/Winter). Leave blank if unknown.'
            ),
            'Has duplicate copy': st.column_config.CheckboxColumn(
                'Has duplicate copy',
                help='Check if the library owns two physical copies of this issue'
            ),
            'Missing / not owned': st.column_config.CheckboxColumn(
                'Missing / not owned',
                help='Check if Hood College does not hold this issue — it will be excluded from the Excel export'
            ),
            'Source': st.column_config.TextColumn('Source', disabled=True),
            'Type': st.column_config.TextColumn('Type', disabled=True),
        }
    )

    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button('Looks good — generate Excel ✓', type='primary'):
            st.session_state.df = edited_df
            go(6)
            st.rerun()
    with col2:
        if st.button('← Back'):
            go(4)
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# STEP 6 — DOWNLOAD
# ═══════════════════════════════════════════════════════════════════════════

elif st.session_state.step == 6:
    show_phase1_progress()
    st.header('Step 6 — Download & Import')

    df = st.session_state.df
    issn = st.session_state.issn
    mms_id = st.session_state.mms_id
    holding_id = st.session_state.holding_id
    title = st.session_state.journal_title or issn

    excel_buf = dataframe_to_excel(df, mms_id, holding_id, issn)

    missing_count = int(df['Missing / not owned'].sum()) if 'Missing / not owned' in df.columns else 0
    exportable_df = df[~df['Missing / not owned']] if 'Missing / not owned' in df.columns else df
    dup_count = int(exportable_df['Has duplicate copy'].sum()) if 'Has duplicate copy' in exportable_df.columns else 0
    total_rows = len(exportable_df) + dup_count
    manual_count = int((df['Source'] == 'Manual').sum()) if 'Source' in df.columns else 0

    st.success(
        f'Your Excel file is ready — **{total_rows} item rows** '
        f'({len(exportable_df)} issues + {dup_count} duplicate copy rows).'
    )

    safe_title = ''.join(c if c.isalnum() or c in ' _-' else '' for c in title)[:30].strip()
    filename = f"items_{issn.replace('-', '')}_{safe_title}.xlsx".replace(' ', '_')

    st.download_button(
        label='⬇  Download Excel file',
        data=excel_buf,
        file_name=filename,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        type='primary',
    )

    st.divider()
    st.subheader('Summary')

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric('Total records', len(df))
    col2.metric('Excluded (missing)', missing_count)
    col3.metric('Exported records', len(exportable_df))
    col4.metric('Duplicate copy rows', dup_count)
    col5.metric('Total Excel rows', total_rows)

    if missing_count > 0:
        st.info(
            f'{missing_count} records were marked as missing / not owned and excluded from the export. '
            'They remain visible in the review table if you need to go back and change anything.'
        )

    if manual_count > 0:
        st.warning(
            f'{manual_count} records were entered manually and could not be verified '
            'automatically. Please double-check these against the physical journal '
            'before importing into Alma.'
        )

    st.divider()
    st.subheader('Next steps in Alma')
    st.markdown(
        '1. Open the downloaded Excel file and verify it looks correct  \n'
        '2. In Alma, go to any page → three dots menu (⋮) → **Cloud App Center**  \n'
        '3. Open **Items Creator by Excel**  \n'
        '4. Upload the file  \n'
        '5. Verify the items appear correctly in Alma'
    )

    st.divider()

    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        if st.button(
            'Continue to Phase 2 — Analyze MARC Pattern →',
            type='primary',
            help='Use the approved items list to generate 853/863/866 MARC holdings fields'
        ):
            go(7)
            st.rerun()
    with col2:
        if st.button('Start over with a new journal'):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
    with col3:
        if st.button('← Back'):
            go(5)
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# STEP 7 — PHASE 2 ENTRY: LOAD ITEMS LIST
# ═══════════════════════════════════════════════════════════════════════════

elif st.session_state.step == 7:
    show_phase2_progress()
    st.header('Step 1 — Load Items List')

    # ── Option A: continuing from Phase 1 ───────────────────────────────────
    if st.session_state.get('all_issues'):
        title = st.session_state.get('journal_title', '') or st.session_state.get('issn', '')
        st.success(
            f'Continuing from Phase 1 — **{len(st.session_state["all_issues"])} item records** '
            f'loaded{(" for *" + title + "*") if title else ""}.'
        )

        col1, col2 = st.columns([2, 1])
        with col1:
            if st.button('Analyze this items list →', type='primary'):
                issues = normalize_session_issues(st.session_state.all_issues)
                st.session_state.p2_issues = issues
                st.session_state.p2_journal_title = st.session_state.get('journal_title', '')
                st.session_state.p2_issn = st.session_state.get('issn', '')
                st.session_state.p2_analysis = None
                go(8)
                st.rerun()
        with col2:
            if st.button('← Back to Phase 1'):
                go(6)
                st.rerun()

        st.divider()
        st.write('Or upload a different items file instead:')

    # ── Option B: upload a file ──────────────────────────────────────────────
    else:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
            <div class="tool-card" style="min-height:130px;">
                <p class="tool-card-title" style="font-size:0.95rem;">Phase 1 export format</p>
                <p class="tool-card-body" style="font-size:0.88rem;">
                Excel file downloaded from this tool's Phase 1 step
                (columns: mms_id, holding_id, enumeration_a, enumeration_b,
                chronology_i, chronology_j, description…).
                </p>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown("""
            <div class="tool-card" style="min-height:130px;">
                <p class="tool-card-title" style="font-size:0.95rem;">Alma export format</p>
                <p class="tool-card-body" style="font-size:0.88rem;">
                Excel file exported directly from Alma's physical items list
                (columns: Barcode, Year, Volume, Description…).
                </p>
            </div>
            """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        'Upload items list (.xlsx)',
        type=['xlsx'],
        label_visibility='collapsed'
    )

    if uploaded_file:
        import zipfile as _zf
        MAX_BYTES = 10 * 1024 * 1024
        if uploaded_file.size > MAX_BYTES:
            st.error(
                f'File is too large ({uploaded_file.size / 1024 / 1024:.1f} MB). '
                'Maximum allowed size is 10 MB.'
            )
        else:
            try:
                uploaded_file.seek(0)
                with _zf.ZipFile(uploaded_file) as z:
                    names = z.namelist()
                if '[Content_Types].xml' not in names or not any(
                    n.startswith('xl/worksheets/') for n in names
                ):
                    raise ValueError('Invalid xlsx structure.')
                uploaded_file.seek(0)

                fmt = detect_format(uploaded_file)
                uploaded_file.seek(0)

                if fmt == 'phase1':
                    issues = parse_phase1_upload(uploaded_file)
                    fmt_label = 'Phase 1 export format'
                elif fmt == 'alma':
                    issues = parse_alma_export(uploaded_file)
                    fmt_label = 'Alma export format'
                else:
                    st.error(
                        'Could not identify the file format. '
                        'Please upload a Phase 1 export or an Alma items export.'
                    )
                    issues = []
                    fmt_label = ''

                if issues:
                    st.success(
                        f'**{uploaded_file.name}** accepted — '
                        f'**{len(issues)} items** detected ({fmt_label}).'
                    )
                    if st.button('Analyze this items list →', type='primary', key='btn_upload_analyze'):
                        st.session_state.p2_issues = issues
                        st.session_state.p2_journal_title = ''
                        st.session_state.p2_issn = ''
                        st.session_state.p2_analysis = None
                        go(8)
                        st.rerun()
                elif fmt_label:
                    st.warning('The file was recognized but no item records could be parsed from it.')

            except (_zf.BadZipFile, ValueError):
                st.error(
                    'This does not appear to be a valid Excel (.xlsx) file. '
                    'Please check the file and try again.'
                )

    if not st.session_state.get('all_issues') and not uploaded_file:
        st.write('')
        if st.button('← Back to Home'):
            go(0)
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# STEP 8 — PHASE 2: PATTERN ANALYSIS RESULTS
# ═══════════════════════════════════════════════════════════════════════════

elif st.session_state.step == 8:
    show_phase2_progress()
    st.header('Step 2 — Publication Pattern Analysis')

    issues = st.session_state.get('p2_issues', [])
    if not issues:
        st.error('No items loaded. Please go back and load an items list.')
        if st.button('← Back'):
            go(7)
            st.rerun()
        st.stop()

    # Run analysis (cache in session state so it doesn't re-run on every interaction)
    if st.session_state.p2_analysis is None:
        with st.spinner('Analyzing publication pattern…'):
            st.session_state.p2_analysis = analyze_pattern(issues)

    analysis = st.session_state.p2_analysis
    sections = analysis.get('sections', [])
    vols = analysis.get('all_vols', [])
    title = st.session_state.get('p2_journal_title', '')

    # ── Data summary ──────────────────────────────────────────────────────
    if title:
        st.subheader(f'"{title}"')

    col1, col2, col3 = st.columns(3)
    col1.metric('Items analyzed', analysis.get('total_regular', 0))
    col2.metric('Volumes', f'{vols[0]}–{vols[-1]}' if vols else '—')
    col3.metric('Sections detected', len(sections))

    special = analysis.get('special_items', [])
    if special:
        from collections import Counter
        types = Counter(i.get('item_type', 'other') for i in special)
        st.caption(
            'Special items excluded from pattern analysis: '
            + ', '.join(f'{n} {t}(s)' for t, n in types.items())
        )

    # ── Warnings ──────────────────────────────────────────────────────────
    for w in analysis.get('warnings', []):
        st.warning(w)

    st.divider()

    # ── Pattern summary ───────────────────────────────────────────────────
    if analysis.get('transition_vol'):
        st.info(
            f'**Transition detected at v.{analysis["transition_vol"]}** — '
            'the publication pattern changes at this point. '
            'Two separate MARC sections will be generated.'
        )

    for i, s in enumerate(sections, 1):
        vol_range = f'v.{s["start_vol"]}–v.{s["end_vol"]}'
        yr_range = f'{s["start_year"] or "?"}–{s["end_year"] or "?"}'
        if s['is_regular']:
            label = f'Section {i}: {vol_range} ({yr_range}) — **Regular**'
            color = 'green'
        else:
            label = f'Section {i}: {vol_range} ({yr_range}) — **Irregular**'
            color = 'orange'

        with st.expander(label, expanded=True):
            if s['is_regular']:
                st.success(
                    f'Frequency: **{s["freq_label"].capitalize()}** — '
                    f'{int(s["consistency"]*100)}% of volumes have {s["count_mode"]} issue(s). '
                    f'Chronology: **{s["chron_type"]}**.  \n'
                    f'→ Will generate **853 + 863**. Alma auto-generates 866 on save.'
                )
            else:
                counts = sorted(set(s['vol_counts'].values()))
                st.warning(
                    f'No consistent frequency — volumes have {counts} issue(s).  \n'
                    f'→ Will generate **866 textual statement** only.'
                )

            if s.get('outlier_vols'):
                st.caption(
                    f'Outlier volumes (different issue count): '
                    + ', '.join(f'v.{v}' for v in s['outlier_vols'][:5])
                    + (' …' if len(s['outlier_vols']) > 5 else '')
                )

    col1, col2 = st.columns([2, 1])
    with col1:
        if sections:
            if st.button('Generate MARC fields →', type='primary'):
                marc = generate_marc_fields(analysis)
                st.session_state.p2_marc_sections = marc
                go(9)
                st.rerun()
        else:
            st.error('No sections could be analyzed. Please check your items list.')
    with col2:
        if st.button('← Back'):
            go(7)
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# STEP 9 — PHASE 2: MARC FIELDS PREVIEW & COPY
# ═══════════════════════════════════════════════════════════════════════════

elif st.session_state.step == 9:
    show_phase2_progress()
    st.header('Step 3 — Review & Copy MARC Fields')

    marc_sections = st.session_state.get('p2_marc_sections', [])
    if not marc_sections:
        st.error('No MARC fields generated. Please go back.')
        if st.button('← Back'):
            go(8)
            st.rerun()
        st.stop()

    st.write(
        'Review the generated MARC fields below. '
        'Each field is shown ready to paste into the Alma Metadata Editor. '
        'You can edit the text if anything needs adjusting before you copy it.'
    )
    st.caption(
        '**How to use:** Open the holdings record in Alma → Metadata Editor → '
        'add each field in the correct position → save. '
        'Alma will auto-generate the 866 field(s) from the 853+863 on save.'
    )

    all_field_texts = []

    for sec in marc_sections:
        st.subheader(sec['label'])

        # Explanation
        st.markdown(sec['explanation'])

        # Editable text area
        default_text = '\n'.join(sec['fields'])
        key = f'marc_edit_{sec["link_num"]}'
        if key not in st.session_state:
            st.session_state[key] = default_text

        edited = st.text_area(
            'MARC fields (editable)',
            value=st.session_state[key],
            height=max(80, 40 * len(sec['fields'])),
            key=f'_ta_{sec["link_num"]}',
            label_visibility='collapsed',
        )
        st.session_state[key] = edited
        all_field_texts.append(edited)

        # Read-only copy block
        st.code(edited, language=None)
        st.divider()

    # Copy-all block
    st.subheader('All fields combined')
    all_text = '\n'.join(all_field_texts)
    st.code(all_text, language=None)
    st.caption('Click the copy icon (top-right of the code block) to copy all fields at once.')

    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button('Continue to report →', type='primary'):
            report = generate_report(
                st.session_state.get('p2_issues', []),
                st.session_state.get('p2_analysis', {}),
                marc_sections,
                journal_title=st.session_state.get('p2_journal_title', ''),
                issn=st.session_state.get('p2_issn', ''),
            )
            st.session_state.p2_report = report
            go(10)
            st.rerun()
    with col2:
        if st.button('← Back'):
            go(8)
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# STEP 10 — PHASE 2: CATALOGING REPORT
# ═══════════════════════════════════════════════════════════════════════════

elif st.session_state.step == 10:
    show_phase2_progress()
    st.header('Step 4 — Cataloging Report')

    report = st.session_state.get('p2_report', '')
    title = st.session_state.get('p2_journal_title', '') or st.session_state.get('p2_issn', '')

    st.success('MARC fields generated. Copy them into Alma and save the holdings record.')

    if title:
        safe = ''.join(c if c.isalnum() or c in ' _-' else '' for c in title)[:30].strip()
    else:
        safe = 'holdings'
    filename = f'MARC_report_{safe}.txt'.replace(' ', '_')

    st.download_button(
        label='⬇  Download cataloging report (.txt)',
        data=report,
        file_name=filename,
        mime='text/plain',
        type='primary',
    )

    st.divider()
    st.subheader('Report')
    st.code(report, language=None)

    st.divider()
    st.subheader('Next steps in Alma')
    st.markdown(
        '1. Open the **holdings record** for this journal in Alma (Metadata Editor)  \n'
        '2. Add the **853 field** from Step 3 above  \n'
        '3. Add the **863 field(s)** from Step 3 above  \n'
        '4. **Save** the holdings record — Alma will automatically generate the 866 field  \n'
        '5. For any **866-only sections** (irregular), add those manually in the same save  \n'
        '6. Verify the 866 textual statement looks correct after saving'
    )

    st.divider()
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button('← Edit MARC fields'):
            go(9)
            st.rerun()
    with col2:
        if st.button('Analyze another journal (Phase 2)'):
            for k in ['p2_issues', 'p2_journal_title', 'p2_issn',
                      'p2_analysis', 'p2_marc_sections', 'p2_report']:
                st.session_state[k] = DEFAULTS[k]
            go(7)
            st.rerun()
    with col3:
        if st.button('⌂  Home'):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
