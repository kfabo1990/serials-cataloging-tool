"""
app.py — Serials Holdings Cataloging Tool (v1: Item Records Generator)
Hood College Library

To run:  streamlit run app.py
"""

import streamlit as st
from helpers import (
    extract_issn, get_journal_info, get_all_works, works_to_issues,
    sort_issues, parse_manual_text, merge_issues,
    issues_to_dataframe, dataframe_to_excel, generate_claude_prompt,
)

# ── PAGE CONFIG ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title='Serials Cataloging Tool',
    page_icon='📚',
    layout='wide'
)

st.title('📚 Serials Holdings Cataloging Tool')
st.caption('Hood College Library — v1: Alma Item Records Generator')

# ── SESSION STATE ────────────────────────────────────────────────────────────
# Session state persists data between Streamlit reruns (every button click reruns the script).

DEFAULTS = {
    'step': 1,
    'issn': None,
    'journal_title': '',
    'crossref_issues': [],
    'manual_text': '',
    'all_issues': [],
    'mms_id': '',
    'holding_id': '',
    'df': None,
    'crossref_done': False,
}

for key, val in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = val


def go(step: int):
    st.session_state.step = step


# ── PROGRESS BAR ─────────────────────────────────────────────────────────────

STEPS = ['ISSN', 'Confirm', 'Discover', 'Alma IDs', 'Review', 'Download']
st.progress((st.session_state.step - 1) / (len(STEPS) - 1))
cols = st.columns(len(STEPS))
for i, label in enumerate(STEPS):
    color = ':blue[**' if i + 1 == st.session_state.step else ''
    end = '**]' if i + 1 == st.session_state.step else ''
    cols[i].markdown(f'{color}{i + 1}. {label}{end}', unsafe_allow_html=False)

st.divider()


# ═══════════════════════════════════════════════════════════════════════════
# STEP 1 — ISSN / URL INPUT
# ═══════════════════════════════════════════════════════════════════════════

if st.session_state.step == 1:
    st.header('Step 1 — Enter ISSN or Journal URL')
    st.write(
        'Type the ISSN directly (e.g. **0046-4813**), or paste any URL for the journal — '
        'the tool will find the ISSN automatically.'
    )

    user_input = st.text_input(
        'ISSN or URL',
        placeholder='0046-4813   or   https://www.tandfonline.com/loi/...',
        label_visibility='collapsed'
    )

    if st.button('Continue →', type='primary'):
        if not user_input.strip():
            st.error('Please enter an ISSN or a URL.')
        else:
            found = extract_issn(user_input.strip())
            if found:
                st.session_state.issn = found
                st.session_state.crossref_done = False  # reset so step 3 fetches fresh
                go(2)
                st.rerun()
            else:
                st.error(
                    'No ISSN found. An ISSN looks like **0046-4813** '
                    '(four digits, a dash, four more digits). '
                    'Please check and try again, or type just the ISSN.'
                )


# ═══════════════════════════════════════════════════════════════════════════
# STEP 2 — CONFIRM ISSN
# ═══════════════════════════════════════════════════════════════════════════

elif st.session_state.step == 2:
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
# STEP 3 — CROSSREF DISCOVERY + GAP HANDLING
# ═══════════════════════════════════════════════════════════════════════════

elif st.session_state.step == 3:
    st.header('Step 3 — Discover Journal Issues')
    issn = st.session_state.issn

    # Only call the API once — results are stored in session state
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

    # ── Show what CrossRef returned ──────────────────────────────────────────
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
                 'Year': i['year'], 'Month': i['month']}
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

    # ── Manual supplement ────────────────────────────────────────────────────
    st.subheader('Add records that CrossRef is missing')
    st.write(
        'If early volumes are missing, or this journal has no electronic version, '
        'you can extract the data from physical journals using Claude.ai and paste it here.'
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
        '(e.g.  v.1 n.1 Apr 1969)',
        value=st.session_state.manual_text,
        height=180,
        placeholder='v.1 n.1 Apr 1969\nv.1 n.2 Jul 1969\nv.1 n.3 Oct 1969\nv.2 n.1 Jan 1970',
    )
    st.session_state.manual_text = manual_text

    # ── Navigation ───────────────────────────────────────────────────────────
    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button('Continue with these records →', type='primary'):
            manual = parse_manual_text(manual_text) if manual_text.strip() else []
            all_issues = sort_issues(merge_issues(issues, manual))

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
            st.info(
                f'You have **{len(issues)} CrossRef records** and '
                f'**{len(manual_preview)} manual records** — '
                f'they will be merged. Manual records take priority if there is overlap.'
            )


# ═══════════════════════════════════════════════════════════════════════════
# STEP 4 — ALMA CONSTANT FIELDS
# ═══════════════════════════════════════════════════════════════════════════

elif st.session_state.step == 4:
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
    st.header('Step 5 — Review & Edit Item Records')

    df = st.session_state.df
    manual_count = int((df['Source'] == 'Manual').sum()) if 'Source' in df.columns else 0
    crossref_count = int((df['Source'] == 'CrossRef').sum()) if 'Source' in df.columns else 0

    st.write(
        f'**{len(df)} issues** ready — {crossref_count} from CrossRef, '
        f'{manual_count} added manually. '
        'Edit any cell directly. Check **Has duplicate copy** for issues where '
        'the library owns two physical copies (a second row will be added automatically).'
    )

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
            'Month': st.column_config.TextColumn('Month', help='3-letter abbreviation e.g. Apr — leave blank if unknown'),
            'Has duplicate copy': st.column_config.CheckboxColumn(
                'Has duplicate copy',
                help='Check if the library owns two physical copies of this issue'
            ),
            'Source': st.column_config.TextColumn('Source', disabled=True),
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
    st.header('Step 6 — Download & Import')

    df = st.session_state.df
    issn = st.session_state.issn
    mms_id = st.session_state.mms_id
    holding_id = st.session_state.holding_id
    title = st.session_state.journal_title or issn

    excel_buf = dataframe_to_excel(df, mms_id, holding_id, issn)

    dup_count = int(df['Has duplicate copy'].sum()) if 'Has duplicate copy' in df.columns else 0
    total_rows = len(df) + dup_count
    manual_count = int((df['Source'] == 'Manual').sum()) if 'Source' in df.columns else 0

    st.success(f'Your Excel file is ready — **{total_rows} item rows** ({dup_count} duplicate copy rows included).')

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

    col1, col2, col3, col4 = st.columns(4)
    col1.metric('Issues found', len(df))
    col2.metric('Duplicate copy rows', dup_count)
    col3.metric('Total Excel rows', total_rows)
    col4.metric('Manual entries', manual_count)

    if manual_count > 0:
        st.warning(
            f'{manual_count} records were entered manually and could not be verified '
            'automatically. Please double-check these against the physical journal '
            'before importing into Alma.'
        )

    st.divider()
    st.subheader('Next steps')
    st.markdown(
        '1. Open the downloaded Excel file and verify it looks correct  \n'
        '2. In Alma, go to any page → three dots menu (⋮) → **Cloud App Center**  \n'
        '3. Open **Items Creator by Excel**  \n'
        '4. Upload the file  \n'
        '5. Verify the items appear correctly in Alma'
    )

    st.divider()
    if st.button('Start over with a new journal'):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
