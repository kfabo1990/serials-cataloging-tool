"""
helpers.py — all data-processing logic (no UI here)
"""

import re
import requests
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

CROSSREF_EMAIL = 'krisztina.fabo@yahoo.com'

MONTH_ABBR = {
    1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr',
    5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Aug',
    9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
}

_MONTH_NAMES = r'Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec'
_SEASON_NAMES = r'Spring|Summer|Fall|Autumn|Winter'

# Matches: single month, combined months (Jan/Feb, Jan-Feb, Mar-Apr),
# single season, combined seasons (Fall/Winter, Spring-Summer)
MONTH_SEASON_RE = re.compile(
    r'\b((?:' + _MONTH_NAMES + r')(?:[/\-](?:' + _MONTH_NAMES + r'))?'
    r'|(?:' + _SEASON_NAMES + r')(?:[/\-](?:' + _SEASON_NAMES + r'))?)\b',
    re.IGNORECASE
)


def _normalize_month_season(s: str) -> str:
    """Title-case each component; normalize separator to /."""
    sep_m = re.search(r'[/\-]', s)
    if sep_m:
        parts = s.split(sep_m.group(), 1)
        return '/'.join(p.strip().capitalize() for p in parts)
    return s.strip().capitalize()


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
    Uses cursor-based pagination — handles journals with thousands of articles.
    Returns a list of raw CrossRef work dicts.
    """
    works = []
    cursor = '*'
    max_pages = 20  # 20 × 1000 = 20,000 articles max

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

        existing = seen.get(key)
        if not existing or (year and (not existing['year'] or year < existing['year'])):
            seen[key] = {
                'volume': vol,
                'issue': iss,
                'year': year,
                'month': month,
                'is_duplicate': False,
                'source': 'CrossRef',
                'missing': False,
                'item_type': 'issue',
            }

    return list(seen.values())


# ── HOLDINGS RANGE FILTER ────────────────────────────────────────────────────

def filter_by_holdings_range(issues: list, start_vol: int, end_vol: int):
    """
    Keep only issues whose volume number falls within [start_vol, end_vol].
    Issues with unparseable volume numbers are kept as-is.
    Does not affect manual entries, indexes, or supplements (handle those separately).
    """
    result = []
    for issue in issues:
        try:
            vol = int(str(issue['volume']).split('-')[0])
        except (ValueError, TypeError):
            result.append(issue)
            continue
        if start_vol <= vol <= end_vol:
            result.append(issue)
    return result


def get_volume_year_map(issues: list) -> dict:
    """Return {vol_int: earliest_year} for all issues with parseable integer volumes."""
    vol_years: dict = {}
    for iss in issues:
        try:
            vol = int(str(iss['volume']).split('-')[0])
        except (ValueError, TypeError):
            continue
        year = iss.get('year')
        if vol not in vol_years or (year and (vol_years[vol] is None or year < vol_years[vol])):
            vol_years[vol] = year
    return vol_years


# ── MANUAL TEXT PARSING ──────────────────────────────────────────────────────

def parse_manual_text(text: str):
    """
    Parse text pasted from a Claude extraction session.
    One issue per line, e.g.:
        v.1 n.1 Apr 1969
        v.1 n.2 Jan/Feb 1970
        v.2 n.1 Spring 1971
        v.2 n.3-4 1972          (combined issue, no month)
        v.3 1973                 (volume only, no issue number)
    """
    issues = []
    for line in text.strip().splitlines():
        line = line.strip()
        if not line:
            continue

        vol_m = re.search(r'v\.?\s*(\d+)', line, re.IGNORECASE)
        iss_m = re.search(r'n(?:o)?\.?\s*([\d]+(?:[-–][\d]+)?)', line, re.IGNORECASE)
        year_m = re.search(r'\b(19\d{2}|20\d{2})\b', line)
        month_m = MONTH_SEASON_RE.search(line)

        if not vol_m and not year_m:
            continue

        issues.append({
            'volume': vol_m.group(1) if vol_m else '',
            'issue': iss_m.group(1) if iss_m else '',
            'year': int(year_m.group(1)) if year_m else None,
            'month': _normalize_month_season(month_m.group(1)) if month_m else '',
            'is_duplicate': False,
            'source': 'Manual',
            'missing': False,
            'item_type': 'issue',
        })

    return issues


# ── MERGE ────────────────────────────────────────────────────────────────────

def merge_issues(crossref_issues: list, manual_issues: list):
    """
    Merge CrossRef and manual records. Manual entries win on (volume, issue) collision —
    the cataloger knows better than the API.
    """
    merged = {(i['volume'], i['issue']): i for i in crossref_issues}
    for issue in manual_issues:
        merged[(issue['volume'], issue['issue'])] = issue
    return list(merged.values())


# ── SORTING ──────────────────────────────────────────────────────────────────

def _sort_key(issue: dict):
    # Regular issues sort before special items (indexes, supplements, etc.)
    item_type = issue.get('item_type', 'issue')
    type_order = 0 if item_type == 'issue' else 1
    try:
        vol = int(str(issue['volume']).split('-')[0])
    except (ValueError, KeyError, TypeError):
        vol = 9999
    try:
        iss = int(str(issue['issue']).split('-')[0])
    except (ValueError, KeyError, TypeError):
        iss = 0
    return (type_order, vol, iss)


def sort_issues(issues: list):
    return sorted(issues, key=_sort_key)


# ── BARCODE & DESCRIPTION ────────────────────────────────────────────────────

def make_barcode(issn: str, vol: str, issue: str, is_copy: bool = False, item_type: str = 'issue'):
    issn_digits = issn.replace('-', '')

    if item_type == 'index':
        # vol is a range like "1-10"; keep the dash for readability
        barcode = f'HCP{issn_digits}v{vol}Index'
    else:
        # Strip characters unsafe in barcodes, but keep digits, letters, and dash
        iss_clean = re.sub(r'[^a-zA-Z0-9\-]', '', issue)
        if vol and iss_clean:
            barcode = f'HCP{issn_digits}v{vol}n{iss_clean}'
        elif vol:
            barcode = f'HCP{issn_digits}v{vol}'
        elif iss_clean:
            barcode = f'HCP{issn_digits}n{iss_clean}'
        else:
            barcode = f'HCP{issn_digits}unknown'

    return barcode + ('copy' if is_copy else '')


def make_description(vol: str, issue: str, year, month: str, is_copy: bool = False, item_type: str = 'issue'):
    date = f'{month}, {year}' if month and year else (str(year) if year else '')

    if item_type == 'index':
        desc = f'Index to v.{vol} ({date})' if date else f'Index to v.{vol}'
    elif item_type in ('supplement', 'part', 'special'):
        label = issue if issue else 'Suppl.'
        if vol:
            desc = f'v. {vol}, {label} ({date})' if date else f'v. {vol}, {label}'
        else:
            desc = f'{label} ({date})' if date else label
    else:
        if vol and issue:
            desc = f'v. {vol}, n. {issue}. ({date})' if date else f'v. {vol}, n. {issue}.'
        elif vol:
            desc = f'v. {vol}. ({date})' if date else f'v. {vol}.'
        elif issue:
            desc = f'n. {issue}. ({date})' if date else f'n. {issue}.'
        else:
            desc = f'({date})' if date else ''

    return ('copy of ' + desc) if is_copy else desc


# ── DATAFRAME ────────────────────────────────────────────────────────────────

def issues_to_dataframe(issues: list):
    """Build the editable review DataFrame shown to the cataloger in Step 5."""
    rows = []
    for iss in issues:
        rows.append({
            'Volume': str(iss.get('volume', '')),
            'Issue': str(iss.get('issue', '')),
            'Year': iss.get('year'),
            'Month/Season': iss.get('month', ''),
            'Has duplicate copy': iss.get('is_duplicate', False),
            'Missing / not owned': iss.get('missing', False),
            'Source': iss.get('source', 'CrossRef'),
            'Type': iss.get('item_type', 'issue'),
        })
    return pd.DataFrame(rows)


# ── EXCEL EXPORT ─────────────────────────────────────────────────────────────

def _sanitize_for_excel(value):
    """
    Prevent formula injection (CSV/Excel injection).
    Excel treats any cell value starting with =, +, -, @, tab, or CR as a formula.
    Prefixing with a single apostrophe neutralizes this — Excel displays the value
    as plain text and does not execute it. Only affects values that start with
    these characters; all normal data (IDs, months, descriptions) passes through unchanged.
    """
    if not isinstance(value, str):
        return value
    if value and value[0] in ('=', '+', '-', '@', '\t', '\r'):
        return "'" + value
    return value


def dataframe_to_excel(df: pd.DataFrame, mms_id: str, holding_id: str, issn: str):
    """
    Convert the approved DataFrame into an Alma Item Import Excel file.
    - Rows marked 'Missing / not owned' are excluded.
    - mms_id and holding_id columns are forced to TEXT format.
    - Index and supplement rows generate appropriate enum fields.
    - All user-sourced string values are sanitized against formula injection.
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
        if bool(row.get('Missing / not owned', False)):
            continue

        vol = str(row['Volume']).strip()
        iss = str(row['Issue']).strip()
        year = row['Year']
        month = str(row.get('Month/Season', '')).strip() if pd.notna(row.get('Month/Season', '')) else ''
        is_dup = bool(row.get('Has duplicate copy', False))
        item_type = str(row.get('Type', 'issue')).strip()

        if not vol and not iss:
            continue

        year_val = year if pd.notna(year) else None

        barcode = make_barcode(issn, vol, iss, item_type=item_type)
        desc = make_description(vol, iss, year_val, month, item_type=item_type)

        if item_type == 'index':
            enum_a = f'v. {vol}' if vol else ''
            enum_b = 'Index'
        elif item_type in ('supplement', 'part', 'special'):
            enum_a = f'v. {vol}' if vol else ''
            enum_b = iss if iss else 'Suppl.'
        else:
            enum_a = f'v. {vol}' if vol else ''
            enum_b = f'n. {iss}' if iss else ''

        data_row = [
            _sanitize_for_excel(mms_id),
            _sanitize_for_excel(holding_id),
            barcode,
            'Journal - Issue',
            'Hood Serials Item',
            enum_a,
            enum_b,
            year_val if year_val is not None else '',
            _sanitize_for_excel(month),
            _sanitize_for_excel(desc),
        ]
        ws.append(data_row)

        # Duplicate copy row — only for regular issues
        if is_dup and item_type == 'issue':
            copy_row = list(data_row)
            copy_row[2] = make_barcode(issn, vol, iss, is_copy=True, item_type=item_type)
            copy_row[9] = _sanitize_for_excel(
                make_description(vol, iss, year_val, month, is_copy=True, item_type=item_type)
            )
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
v.{{VOLUME}} n.{{ISSUE}} {{MONTH-OR-SEASON}} {{YEAR}}

If no month or season is visible, use:
v.{{VOLUME}} n.{{ISSUE}} {{YEAR}}

If the journal has no issue numbers (volume only), use:
v.{{VOLUME}} {{MONTH-OR-SEASON}} {{YEAR}}

For combined months or seasons, use the text as printed, for example:
v.{{VOLUME}} n.{{ISSUE}} Jan/Feb {{YEAR}}
v.{{VOLUME}} n.{{ISSUE}} Spring {{YEAR}}
v.{{VOLUME}} n.{{ISSUE}} Fall/Winter {{YEAR}}
v.{{VOLUME}} n.{{ISSUE}} Jan-Mar {{YEAR}}

For combined issues (e.g., issue 3 and 4 together), use:
v.{{VOLUME}} n.{{FIRST}}-{{LAST}} {{MONTH}} {{YEAR}}

Rules you must follow:
- Do NOT invent or guess any information
- Do NOT fill in months or seasons that are not clearly shown
- If any information is unclear, leave that field blank
- Do NOT include issues marked as missing
- Use 3-letter month abbreviations: Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec
- For seasons use: Spring Summer Fall Winter (or Autumn)

Example of correct output:
v.1 n.1 Apr 1969
v.1 n.2 Jan/Feb 1970
v.2 n.1 Spring 1971
v.2 n.3-4 1972"""
