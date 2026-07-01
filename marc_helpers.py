"""
marc_helpers.py — Phase 2: parse items lists, detect publication patterns,
generate MARC 21 Holdings fields (853 / 863 / 866).
"""

from collections import Counter, defaultdict
import re
import zipfile
from xml.etree import ElementTree as ET
from io import BytesIO

# ── CONSTANTS ────────────────────────────────────────────────────────────────

MONTH_TO_CODE = {
    'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
    'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
    'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12',
}
SEASON_TO_CODE = {
    'spring': '21', 'summer': '22',
    'fall': '23', 'autumn': '23', 'winter': '24',
}
CODE_TO_MONTH = {v: k.capitalize() for k, v in MONTH_TO_CODE.items()}
CODE_TO_SEASON = {'21': 'Spring', '22': 'Summer', '23': 'Fall', '24': 'Winter'}

W_FROM_COUNT = {1: 'a', 2: 'f', 3: 't', 4: 'q', 6: 'b', 12: 'm'}
FREQ_LABEL = {
    'a': 'annual (1×/year)',
    'f': 'semiannual (2×/year)',
    't': 'triannual (3×/year)',
    'q': 'quarterly (4×/year)',
    'b': 'bimonthly (6×/year)',
    'm': 'monthly (12×/year)',
    'z': 'irregular',
}

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

# ── MONTH / SEASON UTILITIES ─────────────────────────────────────────────────

def _parse_period_code(month_str: str) -> str | None:
    """Convert a month/season string to a 2-digit MARC code, or None."""
    if not month_str:
        return None
    s = month_str.strip().lower()
    # Combined like "Jan/Feb" or "Spring/Summer" → use the first part
    first = re.split(r'[/\-]', s)[0].strip()[:6]
    return MONTH_TO_CODE.get(first[:3]) or SEASON_TO_CODE.get(first)


def _is_season(month_str: str) -> bool:
    s = month_str.strip().lower()
    first = re.split(r'[/\-]', s)[0].strip()
    return first in SEASON_TO_CODE


def _code_to_label(code: str) -> str:
    return CODE_TO_SEASON.get(code) or CODE_TO_MONTH.get(code) or code


# ── FILE PARSING ─────────────────────────────────────────────────────────────

def detect_format(file) -> str:
    """
    Detect whether the uploaded file is a Phase 1 export or an Alma export.
    Returns 'phase1', 'alma', or 'unknown'.
    """
    try:
        file.seek(0)
        with zipfile.ZipFile(file) as z:
            xml = z.read('xl/worksheets/sheet1.xml')
        file.seek(0)

        root = ET.fromstring(xml)
        first_row = root.find(f'.//{{{NS}}}row')
        if first_row is None:
            return 'unknown'

        header_cells = []
        for c in first_row.findall(f'{{{NS}}}c'):
            t = c.get('t', '')
            if t == 'inlineStr':
                el = c.find(f'{{{NS}}}is/{{{NS}}}t')
                header_cells.append((el.text or '').lower() if el is not None else '')
            else:
                v = c.find(f'{{{NS}}}v')
                header_cells.append((v.text or '').lower() if v is not None else '')

        headers = ' '.join(header_cells)
        if 'mms_id' in headers or 'enumeration_a' in headers:
            return 'phase1'
        if 'barcode' in headers and 'volume' in headers and 'description' in headers:
            return 'alma'
    except Exception:
        pass
    return 'unknown'


def _read_xlsx_rows(file) -> list[list]:
    """Parse any xlsx using raw XML to handle both inlineStr and normal cells."""
    file.seek(0)
    with zipfile.ZipFile(file) as z:
        # Try to load shared strings (normal xlsx)
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            ss_xml = z.read('xl/sharedStrings.xml')
            ss_root = ET.fromstring(ss_xml)
            for si in ss_root.findall(f'.//{{{NS}}}si'):
                t_els = si.findall(f'.//{{{NS}}}t')
                shared.append(''.join(t.text or '' for t in t_els))

        xml = z.read('xl/worksheets/sheet1.xml')

    root = ET.fromstring(xml)
    rows = []
    for row in root.findall(f'.//{{{NS}}}row'):
        cells = []
        for c in row.findall(f'{{{NS}}}c'):
            t = c.get('t', '')
            if t == 'inlineStr':
                el = c.find(f'{{{NS}}}is/{{{NS}}}t')
                cells.append(el.text if el is not None else '')
            elif t == 's' and shared:
                v = c.find(f'{{{NS}}}v')
                try:
                    cells.append(shared[int(v.text)])
                except (TypeError, IndexError, ValueError):
                    cells.append('')
            else:
                v = c.find(f'{{{NS}}}v')
                cells.append(v.text if v is not None else '')
        rows.append(cells)
    return rows


def parse_phase1_upload(file) -> list[dict]:
    """Parse a Phase 1 Alma-import-format Excel into issue dicts."""
    rows = _read_xlsx_rows(file)
    if not rows:
        return []

    headers = [str(h or '').lower().strip() for h in rows[0]]

    def col(name):
        try:
            return headers.index(name)
        except ValueError:
            return None

    i_enum_a = col('enumeration_a')
    i_enum_b = col('enumeration_b')
    i_chron_i = col('chronology_i')
    i_chron_j = col('chronology_j')
    i_desc = col('description')

    issues = []
    for row in rows[1:]:
        def get(idx):
            if idx is None or idx >= len(row):
                return ''
            return str(row[idx] or '').strip()

        enum_a = get(i_enum_a)   # e.g. "v. 1"
        enum_b = get(i_enum_b)   # e.g. "n. 3-4"
        chron_i = get(i_chron_i) # e.g. "1969"
        chron_j = get(i_chron_j) # e.g. "Apr"
        desc = get(i_desc)

        # Extract volume number
        vol_m = re.search(r'(\d+)', enum_a)
        vol = vol_m.group(1) if vol_m else ''

        # Extract issue number (may be combined like "3-4")
        iss_m = re.search(r'(\d+(?:[-–]\d+)?)', enum_b)
        iss = iss_m.group(1) if iss_m else ''

        # Detect item type from description
        item_type = 'issue'
        if 'index' in desc.lower():
            item_type = 'index'
        elif 'suppl' in desc.lower():
            item_type = 'supplement'
        elif 'copy of' in desc.lower():
            continue  # skip duplicate copy rows

        try:
            year = int(str(chron_i).split('-')[0])
        except (ValueError, TypeError):
            year = None

        if not vol and not year:
            continue

        issues.append({
            'volume': vol,
            'issue': iss,
            'year': year,
            'month': chron_j,
            'item_type': item_type,
            'missing': False,
            'source': 'Phase1Upload',
        })

    return issues


def parse_alma_export(file) -> list[dict]:
    """
    Parse an Alma items export Excel (uses inlineStr cell format).
    Columns: Barcode, Library, Location, ..., Year, Volume, Description, ...
    """
    rows = _read_xlsx_rows(file)
    if not rows:
        return []

    headers = [str(h or '').lower().strip() for h in rows[0]]

    def col(name):
        try:
            return headers.index(name)
        except ValueError:
            return None

    i_year = col('year')
    i_vol = col('volume')
    i_desc = col('description')

    issues = []
    for row in rows[1:]:
        def get(idx):
            if idx is None or idx >= len(row):
                return ''
            return str(row[idx] or '').strip()

        year_raw = get(i_year)    # may be "2008" or "2008-2010"
        vol_raw = get(i_vol)      # e.g. "v. 50"
        desc = get(i_desc)        # e.g. "v. 50, n. 4. (Oct, 2008)"

        # Extract volume
        vol_m = re.search(r'(\d+)', vol_raw)
        vol = vol_m.group(1) if vol_m else ''

        # Extract issue from description: "n. 3-4" or "n. 3"
        iss_m = re.search(r'n\.\s*(\d+(?:[-–]\d+)?)', desc, re.IGNORECASE)
        iss = iss_m.group(1) if iss_m else ''

        # Extract year (take the first year from a range)
        try:
            year = int(str(year_raw).split('-')[0])
        except (ValueError, TypeError):
            year = None

        # Extract month/season from description: "(Oct, 2008)" or "(Spring, 1969)"
        month = ''
        month_m = re.search(
            r'\(([A-Za-z/]+)[\s,]*\d{4}', desc
        )
        if month_m:
            month = month_m.group(1).strip()

        # Detect item type
        item_type = 'issue'
        if 'index' in desc.lower():
            item_type = 'index'
        elif 'suppl' in desc.lower():
            item_type = 'supplement'
        elif 'copy of' in desc.lower():
            continue  # skip duplicate rows

        if not vol and not year:
            continue

        issues.append({
            'volume': vol,
            'issue': iss,
            'year': year,
            'month': month,
            'item_type': item_type,
            'missing': False,
            'source': 'AlmaExport',
        })

    return issues


def normalize_session_issues(all_issues: list) -> list[dict]:
    """Normalize session-state issues (from Phase 1 wizard) to the standard format."""
    result = []
    for iss in all_issues:
        if iss.get('missing', False):
            continue
        result.append({
            'volume': str(iss.get('volume', '')),
            'issue': str(iss.get('issue', '')),
            'year': iss.get('year'),
            'month': iss.get('month', ''),
            'item_type': iss.get('item_type', 'issue'),
            'missing': False,
            'source': 'Phase1Session',
        })
    return result


# ── PATTERN ANALYSIS ─────────────────────────────────────────────────────────

def _parse_vol_int(vol_str: str):
    try:
        return int(str(vol_str).split('-')[0].strip())
    except (ValueError, TypeError):
        return None


def _parse_iss_int(iss_str: str):
    try:
        return int(str(iss_str).split('-')[0].replace('–', '-').strip())
    except (ValueError, TypeError):
        return 0


def _detect_transition(vols: list, counts: dict):
    """
    Find the volume where the publication pattern sustainedly changes.
    Returns the first volume of the NEW pattern, or None if no clear transition.
    """
    if len(vols) < 6:
        return None

    best_vol = None
    best_score = 0

    for split in range(2, len(vols) - 2):
        before = [counts[v] for v in vols[:split]]
        after = [counts[v] for v in vols[split:]]

        mode_b = Counter(before).most_common(1)[0][0]
        mode_a = Counter(after).most_common(1)[0][0]

        if mode_b == mode_a:
            continue

        cons_b = before.count(mode_b) / len(before)
        cons_a = after.count(mode_a) / len(after)
        avg_diff = abs(mode_a - mode_b)

        score = avg_diff * cons_b * cons_a
        if score > best_score and cons_b >= 0.55 and cons_a >= 0.55 and avg_diff >= 1:
            best_score = score
            best_vol = vols[split]

    return best_vol


def _build_section_info(vols: list, by_vol: dict) -> dict:
    """Build a section descriptor for a contiguous range of volumes."""
    all_items = [item for v in vols for item in by_vol[v]]

    counts = [len(by_vol[v]) for v in vols]
    middle = counts[1:-1] if len(counts) > 2 else counts
    if not middle:
        middle = counts

    count_mode, count_n = Counter(middle).most_common(1)[0]
    consistency = count_n / len(middle)
    w_code = W_FROM_COUNT.get(count_mode, 'z')
    is_regular = consistency >= 0.65 and w_code != 'z'

    # Chronology analysis
    codes = [_parse_period_code(i.get('month', '')) for i in all_items]
    codes = [c for c in codes if c]
    uses_seasons = any(_is_season(i.get('month', '')) for i in all_items if i.get('month'))
    uses_months = any(
        not _is_season(i.get('month', '')) and i.get('month', '')
        for i in all_items
    )
    chron_type = 'season' if uses_seasons else ('month' if uses_months else 'none')

    # Start and end period within a volume (for 853 $x and 863 $j)
    start_period = min(codes) if codes else None
    end_period = max(codes) if codes else None

    # $y patterns: detect combined issues
    y_patterns = _detect_y_patterns(all_items, count_mode, chron_type)

    # Year range
    years = [i['year'] for i in all_items if i.get('year')]
    start_year = min(years) if years else None
    end_year = max(years) if years else None

    # Determine $v (numbering continuity)
    # If each volume starts at issue 1, it's r (reset); if issues never reset, it's c
    v_code = 'r'
    iss_nums = []
    for v in vols:
        iss_in_vol = [_parse_iss_int(i.get('issue', '')) for i in by_vol[v]]
        iss_in_vol = [n for n in iss_in_vol if n > 0]
        if iss_in_vol and min(iss_in_vol) > 1:
            iss_nums.append(min(iss_in_vol))
    if len(iss_nums) > len(vols) * 0.5:
        v_code = 'c'

    return {
        'vols': vols,
        'start_vol': vols[0],
        'end_vol': vols[-1],
        'start_year': start_year,
        'end_year': end_year,
        'count_mode': count_mode,
        'w_code': w_code,
        'freq_label': FREQ_LABEL.get(w_code, 'irregular'),
        'is_regular': is_regular,
        'consistency': consistency,
        'chron_type': chron_type,
        'start_period': start_period,
        'end_period': end_period,
        'v_code': v_code,
        'y_patterns': y_patterns,
        'all_items': all_items,
        'vol_counts': {v: len(by_vol[v]) for v in vols},
        'outlier_vols': [v for v in vols if counts[vols.index(v)] != count_mode],
    }


def _detect_y_patterns(items: list, expected_count: int, chron_type: str) -> list:
    """Detect $y regularity patterns (combined months/seasons, omissions)."""
    y = []
    if chron_type == 'none' or expected_count < 2:
        return y

    # Find combined issues (issue field like "3-4")
    combined_months = []
    for item in items:
        iss = str(item.get('issue', ''))
        if re.search(r'\d[-–]\d', iss) and item.get('month'):
            code = _parse_period_code(item['month'])
            if code:
                combined_months.append(code)

    if combined_months:
        cnt = Counter(combined_months)
        # If the same month appears combined in most volumes, add cm code
        for code, n in cnt.most_common():
            if n >= 2:
                y.append(f'cm{code}')

    return y


def analyze_pattern(issues: list) -> dict:
    """
    Main Phase 2 analysis. Returns a structured dict describing the
    publication pattern and all information needed to generate MARC fields.
    """
    regular = [i for i in issues if i.get('item_type', 'issue') == 'issue']
    special = [i for i in issues if i.get('item_type', 'issue') != 'issue']

    warnings = []

    if not regular:
        return {
            'sections': [],
            'special_items': special,
            'warnings': ['No regular issues found in the items list.'],
            'total': len(issues),
        }

    # Parse and sort
    parsed = []
    for iss in regular:
        vol_int = _parse_vol_int(iss.get('volume', ''))
        parsed.append({**iss, 'vol_int': vol_int})

    parsed.sort(key=lambda x: (x['vol_int'] if x['vol_int'] is not None else 9999,
                                _parse_iss_int(x.get('issue', ''))))

    # Group by volume
    by_vol = defaultdict(list)
    no_vol = []
    for item in parsed:
        if item['vol_int'] is not None:
            by_vol[item['vol_int']].append(item)
        else:
            no_vol.append(item)

    if no_vol:
        warnings.append(
            f'{len(no_vol)} item(s) have no recognizable volume number and were excluded from pattern analysis.'
        )

    vols = sorted(by_vol.keys())

    if not vols:
        return {
            'sections': [],
            'special_items': special,
            'warnings': warnings + ['Could not identify volume numbers in this items list.'],
            'total': len(issues),
        }

    # Count issues per volume
    counts = {v: len(by_vol[v]) for v in vols}

    # Check for outlier first/last volumes
    if len(vols) >= 3:
        middle_mode = Counter([counts[v] for v in vols[1:-1]]).most_common(1)[0][0]
        if counts[vols[0]] < middle_mode:
            warnings.append(
                f'v.{vols[0]} has {counts[vols[0]]} issue(s) (expected {middle_mode}) — '
                'may be a partial first volume.'
            )
        if counts[vols[-1]] < middle_mode:
            warnings.append(
                f'v.{vols[-1]} has {counts[vols[-1]]} issue(s) (expected {middle_mode}) — '
                'may be a partial last volume.'
            )

    # Detect transition
    transition_vol = _detect_transition(vols, counts)

    # Build sections
    if transition_vol and transition_vol in vols:
        split = vols.index(transition_vol)
        section_vol_groups = [vols[:split], vols[split:]]
    else:
        section_vol_groups = [vols]

    sections = [_build_section_info(grp, by_vol) for grp in section_vol_groups if grp]

    # Assign MARC strategy to each section
    for s in sections:
        if s['is_regular']:
            s['marc_type'] = '853_863'
        else:
            s['marc_type'] = '866'

    return {
        'sections': sections,
        'special_items': special,
        'transition_vol': transition_vol,
        'all_vols': vols,
        'total': len(issues),
        'total_regular': len(regular),
        'warnings': warnings,
    }


# ── MARC FIELD GENERATION ────────────────────────────────────────────────────

def _fmt(items: list) -> str:
    """Format a list of '$x value' subfield parts into a spaced string."""
    return '  '.join(items)


def generate_853(section: dict, link_num: int) -> str:
    parts = [f'$8 {link_num}']
    parts.append('$a v.')
    if section['count_mode'] > 1:
        parts.append('$b no.')
    parts.append(f'$u {section["count_mode"]}')
    parts.append(f'$v {section["v_code"]}')
    parts.append('$i (year)')
    if section['chron_type'] == 'month':
        parts.append('$j (month)')
    elif section['chron_type'] == 'season':
        parts.append('$j (season)')
    parts.append(f'$w {section["w_code"]}')
    if section.get('start_period'):
        parts.append(f'$x {section["start_period"]}')
    for yp in section.get('y_patterns', []):
        parts.append(f'$y {yp}')
    return f'853 2 3  {_fmt(parts)}'


def generate_863(section: dict, link_num: int, seq_num: int) -> str:
    parts = [f'$8 {link_num}.{seq_num}']
    parts.append(f'$a {section["start_vol"]}-{section["end_vol"]}')
    if section['count_mode'] > 1:
        parts.append(f'$b 1-{section["count_mode"]}')
    if section['start_year'] and section['end_year']:
        parts.append(f'$i {section["start_year"]}-{section["end_year"]}')
    elif section['start_year']:
        parts.append(f'$i {section["start_year"]}')
    if section['chron_type'] != 'none' and section.get('start_period') and section.get('end_period'):
        parts.append(f'$j {section["start_period"]}-{section["end_period"]}')
    return f'863 3 2  {_fmt(parts)}'


def _format_issue_cite(item: dict) -> str:
    """Format a single issue as a MARC 866 citation: v.1:no.2(1969:Apr.)"""
    vol = item.get('volume', '')
    iss = item.get('issue', '')
    year = item.get('year', '')
    month = item.get('month', '')

    vol_part = f'v.{vol}' if vol else ''
    iss_part = f':no.{iss}' if iss else ''

    date_inner = []
    if year:
        date_inner.append(str(year))
    if month:
        date_inner.append(month)
    date_part = f'({":".join(date_inner)})' if date_inner else ''

    return vol_part + iss_part + date_part


def generate_866(section: dict, link_num: int, seq_num: int) -> str:
    items = sorted(
        section['all_items'],
        key=lambda x: (_parse_vol_int(x.get('volume', '')) or 9999,
                       _parse_iss_int(x.get('issue', '')))
    )
    if not items:
        return ''
    first = _format_issue_cite(items[0])
    last = _format_issue_cite(items[-1])
    stmt = f'{first}-{last}' if first != last else first
    return f'866 4 1  $8 {link_num}.{seq_num}  $a {stmt}'


def generate_marc_fields(analysis: dict) -> list[dict]:
    """
    Generate all MARC fields for the complete analysis.
    Returns a list of section dicts, each containing:
      - label: human-readable section name
      - marc_type: '853_863' or '866'
      - fields: list of MARC field strings (ready to paste into Alma)
      - explanation: plain English reason for these fields
    """
    results = []
    link_num = 1

    for i, section in enumerate(analysis.get('sections', [])):
        label = (
            f'v.{section["start_vol"]}–v.{section["end_vol"]}'
            f' ({section["start_year"] or "?"}–{section["end_year"] or "?"})'
        )
        fields = []

        if section['marc_type'] == '853_863':
            fields.append(generate_853(section, link_num))
            fields.append(generate_863(section, link_num, 1))
            explanation = _explain_853_863(section)
        else:
            fields.append(generate_866(section, link_num, 1))
            explanation = _explain_866(section)

        results.append({
            'label': label,
            'marc_type': section['marc_type'],
            'fields': fields,
            'explanation': explanation,
            'section': section,
            'link_num': link_num,
        })
        link_num += 1

    return results


# ── PLAIN ENGLISH EXPLANATIONS ───────────────────────────────────────────────

def _explain_853_863(section: dict) -> str:
    w = section['w_code']
    freq = FREQ_LABEL.get(w, 'unknown frequency')
    chron = section['chron_type']
    cons_pct = int(section['consistency'] * 100)

    lines = [
        f'**Pattern: {freq.capitalize()}** — {cons_pct}% of volumes '
        f'have {section["count_mode"]} issue(s).',
    ]

    if chron == 'season':
        start = _code_to_label(section.get('start_period', ''))
        end = _code_to_label(section.get('end_period', ''))
        lines.append(f'Uses **seasonal notation** ({start} through {end}).')
    elif chron == 'month':
        start = _code_to_label(section.get('start_period', ''))
        end = _code_to_label(section.get('end_period', ''))
        lines.append(f'Uses **monthly notation** ({start} through {end}).')
    else:
        lines.append('No month/season data available — $j omitted from both fields.')

    for y in section.get('y_patterns', []):
        lines.append(f'Regularity note: `$y {y}` added (combined issues detected).')

    lines.append(
        'Strategy: **853 + 863** — Alma will auto-generate the 866 textual '
        'statement when you save the holdings record.'
    )
    return '  \n'.join(lines)


def _explain_866(section: dict) -> str:
    counts = list(section['vol_counts'].values())
    unique_counts = set(counts)
    lines = [
        f'**Pattern: Irregular** — volumes in this section have '
        f'{sorted(unique_counts)} issue(s) per volume with no consistent frequency.',
        'Strategy: **866 only** — a textual holdings statement is the only option '
        'when there is no regular pattern for Alma to encode. '
        'Alma cannot auto-generate a meaningful 866 from irregular data.',
    ]
    return '  \n'.join(lines)


# ── REPORT ───────────────────────────────────────────────────────────────────

def generate_report(issues: list, analysis: dict, marc_sections: list,
                    journal_title: str = '', issn: str = '') -> str:
    """Generate a plain-text cataloging report."""
    lines = ['MARC HOLDINGS CATALOGING REPORT', '=' * 40]

    if journal_title:
        lines.append(f'Journal: {journal_title}')
    if issn:
        lines.append(f'ISSN:    {issn}')

    vols = analysis.get('all_vols', [])
    total = analysis.get('total_regular', 0)
    special = analysis.get('special_items', [])

    if vols:
        lines.append(f'Volumes: v.{vols[0]}–v.{vols[-1]} ({len(vols)} volumes)')
    lines.append(f'Regular issues analyzed: {total}')
    if special:
        types = Counter(i.get('item_type', 'other') for i in special)
        lines.append(f'Special items (excluded from pattern): '
                     + ', '.join(f'{v} {k}(s)' for k, v in types.items()))

    lines.append('')
    lines.append('PATTERN ANALYSIS')
    lines.append('-' * 40)

    if analysis.get('transition_vol'):
        lines.append(
            f'Transition detected at v.{analysis["transition_vol"]}: '
            'pattern changes at this volume.'
        )

    for i, s in enumerate(analysis.get('sections', []), 1):
        lines.append(
            f'Section {i}: v.{s["start_vol"]}–v.{s["end_vol"]} '
            f'({s["start_year"] or "?"}–{s["end_year"] or "?"}) — '
            + ('REGULAR, ' + FREQ_LABEL.get(s["w_code"], '') if s["is_regular"] else 'IRREGULAR')
        )

    if analysis.get('warnings'):
        lines.append('')
        lines.append('WARNINGS')
        lines.append('-' * 40)
        for w in analysis['warnings']:
            lines.append(f'• {w}')

    lines.append('')
    lines.append('GENERATED MARC FIELDS')
    lines.append('-' * 40)
    for sec in marc_sections:
        lines.append(f'[{sec["label"]}]')
        for field in sec['fields']:
            lines.append(field)
        lines.append('')

    lines.append('=' * 40)
    lines.append('Fields generated by Serials Holdings Cataloging Tool — Hood College Library')
    lines.append('Review all fields before pasting into Alma. Alma will generate 866')
    lines.append('automatically on save for any section with 853+863 fields.')

    return '\n'.join(lines)
