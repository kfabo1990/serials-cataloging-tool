"""
helpers.py — all data-processing logic (no UI here)
"""

import re
import requests
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

CROSSREF_EMAIL = 'krisztina.fabo@yahoo.com'  # "polite pool" gives better CrossRef rate limits

MONTH_ABBR = {
    1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr',
    5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Aug',
    9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
}


# ── ISSN ────────────────────────────────────────────────────────────────────

def extract_issn(text: str):
    """Find the first ISSN pattern (XXXX-XXXX) anywhere in a string or URL."""
    match = re.search(r'\b(\d{4}-\d{3}[\dXx])\b', text)
    return match.group(1).upper() if match else None


# ── CROSSREF ─────────────────────────────────────────────────────────────────

def get_journal_info(issn: str):
    """Return basic journal metadata from CrossRef, or None if not found."""
    try:
        r = requests.get(
            f'https://api.crossref.org/journals/{issn}',
            params={'mailto': CROSSREF_EMAIL},
            timeout=10
        )
        if r.status_code == 200:
            return r.json()['message']
    except requests.RequestException:
        pass
    return None


def get_all_works(issn: str):
    """
    Fetch every article registered in CrossRef for this journal.
    Uses cursor-based pagination so it handles journals with thousands of articles.
    Returns a list of raw CrossRef work dicts.
    """
    works = []
    cursor = '*'
    max_pages = 20  # safety cap: 20 × 1000 = 20,000 articles max

    for _ in range(max_pages):
        try:
            r = requests.get(
                f'https://api.crossref.org/journals/{issn}/works',
                params={
                    'mailto': CROSSREF_EMAIL,
                    'rows': 1000,
                    'cursor': cursor,
                    'select': 'volume,issue,published',
                },
                timeout=30
            )
        except requests.RequestException:
            break

        if r.status_code != 200:
            break

        data = r.json()['message']
        items = data.get('items', [])
        if not items:
            break

        works.extend(items)
        next_cursor = data.get('next-cursor')

        if not next_cursor or len(items) < 1000:
            break
        cursor = next_cursor

    return works


def works_to_issues(works: list):
    """
    Collapse article-level CrossRef records into unique (volume, issue) pairs.
    Keeps the earliest date found for each issue.
    """
    seen = {}
    for work in works:
        vol = str(work.get('volume', '')).strip()
        iss = str(work.get('issue', '')).strip()
        if not vol:
            continue

        key = (vol, iss)
        date_parts = work.get('published', {}).get('date-parts', [[]])[0]
        year = date_parts[0] if len(date_parts) > 0 else None
        month_num = date_parts[1] if len(date_parts) > 1 else None
        month = MONTH_ABBR.get(month_num, '') if month_num else ''

        # Keep earliest year seen for this (vol, issue)
        existing = seen.get(key)
        if not existing or (year and (not existing['year'] or year < existing['year'])):
            seen[key] = {
                'volume': vol,
                'issue': iss,
                'year': year,
                'month': month,
                'is_duplicate': False,
                'source': 'CrossRef',
            }

    return list(seen.values())


# ── MANUAL TEXT PARSING ──────────────────────────────────────────────────────

def parse_manual_text(text: str):
    """
    Parse text pasted from a Claude extraction session.
    Expected format (one line per issue):
        v.1 n.1 Apr 1969
        v.1 n.2 Jul 1969
        v.2 n.3-4 1971        (combined issue, no month)
    Returns list of issue dicts.
    """
    issues = []
    for line in text.strip().splitlines():
        line = line.strip()
        if not line:
            continue

        vol_m = re.search(r'v\.?\s*(\d+)', line, re.IGNORECASE)
        iss_m = re.search(r'n(?:o)?\.?\s*([\d]+(?:-[\d]+)?)', line, re.IGNORECASE)
        year_m = re.search(r'\b(19\d{2}|20\d{2})\b', line)
        month_m = re.search(
            r'\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b',
            line, re.IGNORECASE
        )

        # Skip lines with no volume or year — likely headers/blank
        if not vol_m and not year_m:
            continue

        issues.append({
            'volume': vol_m.group(1) if vol_m else '',
            'issue': iss_m.group(1) if iss_m else '',
            'year': int(year_m.group(1)) if year_m else None,
            'month': month_m.group(1).capitalize() if month_m else '',
            'is_duplicate': False,
            'source': 'Manual',
        })

    return issues


# ── MERGE ────────────────────────────────────────────────────────────────────

def merge_issues(crossref_issues: list, manual_issues: list):
    """
    Merge CrossRef and manual records. Manual entries win if there is a
    (volume, issue) collision — the cataloger knows better than the API.
    """
    merged = {(i['volume'], i['issue']): i for i in crossref_issues}
    for issue in manual_issues:
        merged[(issue['volume'], issue['issue'])] = issue
    return list(merged.values())


# ── SORTING ──────────────────────────────────────────────────────────────────

def _sort_key(issue: dict):
    try:
        vol = int(issue['volume'])
    except (ValueError, KeyError):
        vol = 0
    try:
        iss = int(str(issue['issue']).split('-')[0])
    except (ValueError, KeyError):
        iss = 0
    return (vol, iss)


def sort_issues(issues: list):
    return sorted(issues, key=_sort_key)


# ── BARCODE & DESCRIPTION ────────────────────────────────────────────────────

def make_barcode(issn: str, vol: str, issue: str, is_copy: bool = False):
    issn_digits = issn.replace('-', '')
    barcode = f'HCP{issn_digits}v{vol}n{issue}'
    return barcode + ('copy' if is_copy else '')


def make_description(vol: str, issue: str, year, month: str, is_copy: bool = False):
    if month and year:
        date = f'{month}, {year}'
    elif year:
        date = str(year)
    else:
        date = ''
    desc = f'v. {vol}, n. {issue}. ({date})'
    return ('copy of ' + desc) if is_copy else desc


# ── DATAFRAME ────────────────────────────────────────────────────────────────

def issues_to_dataframe(issues: list):
    """Build the editable review DataFrame shown to the cataloger."""
    rows = []
    for iss in issues:
        rows.append({
            'Volume': str(iss.get('volume', '')),
            'Issue': str(iss.get('issue', '')),
            'Year': iss.get('year'),
            'Month': iss.get('month', ''),
            'Has duplicate copy': iss.get('is_duplicate', False),
            'Source': iss.get('source', 'CrossRef'),
        })
    return pd.DataFrame(rows)


# ── EXCEL EXPORT ─────────────────────────────────────────────────────────────

def dataframe_to_excel(df: pd.DataFrame, mms_id: str, holding_id: str, issn: str):
    """
    Convert the approved DataFrame into an Alma Item Import Excel file.
    mms_id and holding_id columns are forced to TEXT format so Alma
    doesn't misread them as numbers.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    headers = [
        'mms_id', 'holding_id', 'barcode', 'material_type', 'item_policy',
        'enumeration_a', 'enumeration_b', 'chronology_i', 'chronology_j', 'description'
    ]
    ws.append(headers)

    for _, row in df.iterrows():
        vol = str(row['Volume']).strip()
        iss = str(row['Issue']).strip()
        year = row['Year']
        month = str(row['Month']).strip() if pd.notna(row['Month']) else ''
        is_dup = bool(row.get('Has duplicate copy', False))

        if not vol:
            continue

        barcode = make_barcode(issn, vol, iss)
        desc = make_description(vol, iss, year, month)

        data_row = [
            mms_id,
            holding_id,
            barcode,
            'Journal - Issue',
            'Hood Serials Item',
            f'v. {vol}',
            f'n. {iss}',
            year if pd.notna(year) else '',
            month,
            desc,
        ]
        ws.append(data_row)

        if is_dup:
            copy_row = list(data_row)
            copy_row[2] = make_barcode(issn, vol, iss, is_copy=True)
            copy_row[9] = make_description(vol, iss, year, month, is_copy=True)
            ws.append(copy_row)

    # Force mms_id (col A) and holding_id (col B) to text — critical for Alma import
    for excel_row in ws.iter_rows(min_row=2, max_col=2):
        for cell in excel_row:
            cell.number_format = '@'

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── CLAUDE PROMPT ────────────────────────────────────────────────────────────

def generate_claude_prompt(issn: str, journal_title: str = ''):
    title_part = f'for the journal "{journal_title}" (ISSN {issn})' if journal_title else f'(ISSN {issn})'
    return f"""I need to catalog the physical journal holdings {title_part}.

Please extract all volume, issue, and date information from the image or file I am providing.

List every issue on its own line in this exact format:
v.{{VOLUME}} n.{{ISSUE}} {{3-LETTER-MONTH}} {{YEAR}}

If no month is visible, use:
v.{{VOLUME}} n.{{ISSUE}} {{YEAR}}

For combined issues (e.g., issue 3 and 4 together), use:
v.{{VOLUME}} n.{{FIRST}}-{{LAST}} {{MONTH}} {{YEAR}}

Rules you must follow:
- Do NOT invent or guess any information
- Do NOT fill in months that are not clearly shown
- If any information is unclear, leave that field blank
- Do NOT include issues marked as missing
- Use 3-letter month abbreviations: Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec

Example of correct output:
v.1 n.1 Apr 1969
v.1 n.2 Jul 1969
v.1 n.3 Oct 1969
v.2 n.1 Jan 1970
v.2 n.3-4 1971"""
