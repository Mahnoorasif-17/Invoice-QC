"""
PDF Extractor — replaces the Claude API call from the original tool.
Uses pdfplumber to extract text + tables locally (no network, no keys).

Returns the same JSON shape the API used to return:
{
    "groupId": str,
    "groupName": str,
    "invoiceAmount": float,
    "currentPeriodAmount": float,
    "adjustmentChargesTotal": float,
    "adjustmentCreditsTotal": float,
    "currentRoster":          [{firstName, lastName, tier, planCode, cost}],
    "adjustmentChargeRoster": [{firstName, lastName, tier, planCode, coverageMonth, cost}],
    "adjustmentCreditRoster": [{firstName, lastName, tier, planCode, coverageMonth, cost}],
    "_raw_text": str   # for debugging — shown in the "Debug" expander in the UI
}

⚠️  PDF layouts vary. The regex patterns and table-classification rules below are
   best-effort defaults that cover common Redirect Health invoice layouts.
   If your invoice format differs, tune the patterns in:
     • _find_group_id / _find_group_name / _find_amount   (header fields)
     • _classify_table                                    (which roster a table belongs to)
     • _parse_roster_table                                (column mapping)
   The raw extracted text is always returned in `_raw_text` so you can see exactly
   what pdfplumber pulled and adjust accordingly.
"""

import re
from typing import Optional, Any
import pdfplumber


# ─── PUBLIC API ─────────────────────────────────────────────────────
def extract_pdf(file_obj) -> dict:
    """Extract structured invoice data from a PDF file-like object."""
    all_text_parts = []
    all_tables = []

    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ''
            all_text_parts.append(txt)
            for t in (page.extract_tables() or []):
                if t and len(t) > 0:
                    all_tables.append(t)

    text = '\n'.join(all_text_parts)

    return {
        'groupId':                  _find_group_id(text) or '',
        'groupName':                _find_group_name(text) or '',
        'invoiceAmount':            _find_amount(text, [
                                        'invoice total', 'total due', 'grand total',
                                        'amount due', 'total amount', 'total invoice'
                                    ]) or 0,
        'currentPeriodAmount':      _find_amount(text, [
                                        'current period', 'current charges',
                                        'monthly charges', 'current month'
                                    ]) or 0,
        'adjustmentChargesTotal':   _find_amount(text, [
                                        'adjustment charges total', 'total adjustment charges',
                                        'adjustment charges'
                                    ]) or 0,
        'adjustmentCreditsTotal':   _find_amount(text, [
                                        'adjustment credits total', 'total adjustment credits',
                                        'adjustment credits'
                                    ]) or 0,
        'currentRoster':            _collect_roster(all_tables, 'current'),
        'adjustmentChargeRoster':   _collect_roster(all_tables, 'adj_charge'),
        'adjustmentCreditRoster':   _collect_roster(all_tables, 'adj_credit'),
        '_raw_text':                text,
    }


# ─── HEADER FIELD EXTRACTION ────────────────────────────────────────
def _find_group_id(text: str) -> Optional[str]:
    patterns = [
        r'Group\s*(?:ID|#|Number|No\.?)\s*:?\s*([A-Z0-9][A-Z0-9\-]{2,})',
        r'Account\s*(?:ID|#|Number)\s*:?\s*([A-Z0-9][A-Z0-9\-]{2,})',
        r'Customer\s*(?:ID|#|Number)\s*:?\s*([A-Z0-9][A-Z0-9\-]{2,})',
        r'Identifier\s*:?\s*([A-Z0-9][A-Z0-9\-]{2,})',
    ]
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return None


def _find_group_name(text: str) -> Optional[str]:
    patterns = [
        r'Group\s*Name\s*:?\s*([^\n\r]{2,80})',
        r'(?:Bill|Sold)\s*To\s*:?\s*([^\n\r]{2,80})',
        r'Company\s*Name\s*:?\s*([^\n\r]{2,80})',
        r'Customer\s*Name\s*:?\s*([^\n\r]{2,80})',
    ]
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            name = m.group(1).strip()
            # Cut off at common separators that bleed into the next field
            for sep in ['  ', '\t', 'Group ID', 'Group #', 'Invoice', 'Date']:
                if sep in name:
                    name = name.split(sep)[0].strip()
            return name if name else None
    return None


def _find_amount(text: str, keywords: list) -> Optional[float]:
    """Find a dollar amount that follows one of the given keywords."""
    for kw in keywords:
        # Look for: "Keyword [:] $1,234.56" possibly with whitespace/newlines
        pattern = re.escape(kw) + r'\s*[:.]?\s*\$?\s*\(?(-?[\d,]+(?:\.\d{1,2})?)\)?'
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            try:
                val = float(m.group(1).replace(',', ''))
                # Parenthesized amount = negative
                start, end = m.span()
                ctx = text[max(0, start - 1):end + 1]
                if '(' in ctx and ')' in ctx:
                    val = -abs(val)
                return val
            except ValueError:
                continue
    return None


# ─── TABLE CLASSIFICATION + ROSTER PARSING ──────────────────────────
def _collect_roster(tables: list, category: str) -> list:
    rows = []
    for t in tables:
        if _classify_table(t) == category:
            rows.extend(_parse_roster_table(t, category))
    return rows


def _classify_table(table: list) -> Optional[str]:
    """
    Decide whether a table is the current roster, adjustment charges,
    adjustment credits, or none of the above.
    """
    if not table or len(table) < 2:
        return None

    # Look at the first 3 rows of text combined
    blob = ' '.join(
        ' '.join(str(c or '') for c in row) for row in table[:3]
    ).lower()

    has_name_cols = any(k in blob for k in ('first', 'last', 'name'))
    has_amount_col = any(k in blob for k in ('cost', 'amount', 'charge', 'premium', 'total'))

    if not (has_name_cols and has_amount_col):
        return None

    has_coverage_month = ('coverage month' in blob) or ('cov month' in blob) \
        or bool(re.search(r'\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b[^\n]*\d{4}', blob))
    is_adjustment = 'adjustment' in blob or has_coverage_month

    if is_adjustment:
        # Credit vs charge: look for "credit" keyword or negative/parenthesized amounts
        if 'credit' in blob or 'refund' in blob:
            return 'adj_credit'
        # Check for negative amounts in body rows
        for row in table[1:]:
            for cell in row:
                s = str(cell or '')
                if '(' in s and ')' in s and re.search(r'\d', s):
                    return 'adj_credit'
                if s.strip().startswith('-') and re.search(r'\d', s):
                    return 'adj_credit'
        return 'adj_charge'

    return 'current'


def _parse_roster_table(table: list, category: str) -> list:
    """Generic roster table parser. Maps columns by header text."""
    if not table or len(table) < 2:
        return []

    # Find header row = first row with any text
    header_idx = 0
    for i, row in enumerate(table):
        if any(c and str(c).strip() for c in row):
            header_idx = i
            break
    header = [str(c or '').strip().lower() for c in table[header_idx]]

    col_map: dict = {}
    for i, h in enumerate(header):
        if 'first' in h and 'name' in h:
            col_map['firstName'] = i
        elif 'last' in h and 'name' in h:
            col_map['lastName'] = i
        elif h in ('name', 'member name', 'full name', 'employee name', 'member'):
            col_map.setdefault('fullName', i)
        elif 'tier' in h:
            col_map['tier'] = i
        elif ('plan' in h and ('code' in h or 'sku' in h)) or h == 'sku':
            col_map['planCode'] = i
        elif h == 'plan':
            col_map.setdefault('planCode', i)
        elif 'coverage' in h or ('month' in h and 'year' not in h):
            col_map['coverageMonth'] = i
        elif any(k in h for k in ('cost', 'amount', 'charge', 'credit', 'premium', 'total')):
            col_map.setdefault('cost', i)

    out = []
    for row in table[header_idx + 1:]:
        if not any(c and str(c).strip() for c in row):
            continue

        first_cell = str(row[0] or '').strip().lower()
        if first_cell in ('total', 'subtotal', 'grand total', 'sum', 'totals'):
            continue

        def get(key, default=''):
            idx = col_map.get(key)
            if idx is None or idx >= len(row):
                return default
            v = row[idx]
            return str(v or '').strip()

        first = get('firstName')
        last = get('lastName')

        # Split a combined "full name" cell when no first/last columns exist
        if not first and not last and 'fullName' in col_map:
            full = get('fullName')
            if ',' in full:                       # "Last, First M"
                parts = [p.strip() for p in full.split(',', 1)]
                last = parts[0]
                first = parts[1] if len(parts) > 1 else ''
            else:                                  # "First Last"
                parts = full.split(None, 1)
                first = parts[0] if parts else ''
                last = parts[1] if len(parts) > 1 else ''

        if not first and not last:
            continue

        cost_raw = get('cost', '0')
        cost_clean = cost_raw.replace('$', '').replace(',', '').strip()
        if cost_clean.startswith('(') and cost_clean.endswith(')'):
            cost_clean = '-' + cost_clean[1:-1]
        try:
            cost = float(cost_clean) if cost_clean else 0.0
        except ValueError:
            cost = 0.0

        entry = {
            'firstName': first,
            'lastName':  last,
            'tier':      get('tier'),
            'planCode':  get('planCode'),
            'cost':      cost,
        }
        if category in ('adj_charge', 'adj_credit'):
            entry['coverageMonth'] = get('coverageMonth')

        out.append(entry)

    return out
