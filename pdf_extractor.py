"""
PDF Extractor for Redirect Health invoices.

Uses TEXT-line parsing (not pdfplumber's table extraction, which strips
columns when rule lines don't match text positions on this PDF format).

Strategy:
  1. Read ALL pages with pdfplumber, concatenate text
  2. Find header fields with regex (GROUP ID, INVOICE AMOUNT, etc.)
  3. Find roster sections by their section headers:
       CURRENT PERIOD - ROSTER
       ADJUSTMENTS - CHARGES ROSTER
       ADJUSTMENTS - CREDITS ROSTER
  4. Within each section, parse member rows line-by-line
  5. Skip Division:, Total:, and column header lines

Output JSON shape matches what the QC engine expects.
"""

import os
import re
from typing import Optional, Any
import pdfplumber


# ─── PUBLIC API ─────────────────────────────────────────────────────
def extract_pdf(file_obj) -> dict:
    """Extract structured invoice data from a PDF file-like object."""
    # Try to get a filename for group-name fallback
    filename = getattr(file_obj, 'name', '') or ''

    # Read full text from ALL pages
    all_pages_text = []
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            all_pages_text.append(page.extract_text() or '')
    text = '\n'.join(all_pages_text)

    return {
        'groupId':                _find_group_id(text) or '',
        'groupName':              _find_group_name(text, filename) or '',
        'invoiceAmount':          _find_invoice_amount(text) or 0,
        'currentPeriodAmount':    _find_current_period_amount(text) or 0,
        'adjustmentChargesTotal': _find_adj_total(text, 'Charges') or 0,
        'adjustmentCreditsTotal': _find_adj_total(text, 'Credits') or 0,
        'currentRoster':          _parse_roster(text, 'CURRENT PERIOD - ROSTER',
                                                has_coverage_month=False),
        'adjustmentChargeRoster': _parse_roster(text, 'ADJUSTMENTS - CHARGES ROSTER',
                                                has_coverage_month=True),
        'adjustmentCreditRoster': _parse_roster(text, 'ADJUSTMENTS - CREDITS ROSTER',
                                                has_coverage_month=True),
        '_raw_text': text,
        'cobraMembers': _parse_cobra_members(text),
    }


# ─── HEADER FIELD EXTRACTION ────────────────────────────────────────
def _find_group_id(text: str) -> Optional[str]:
    m = re.search(r'GROUP\s*ID\s*:\s*(\S+)', text, re.IGNORECASE)
    return m.group(1).strip() if m else None


def _find_group_name(text: str, filename: str = '') -> Optional[str]:
    """
    Try three strategies in order:
      1. Look for an "Attn Company Broker" header (Redirect Health format)
         and pull the middle column.
      2. Look for explicit "Group Name:" / "Bill To:" labels.
      3. Fall back to parsing the filename (e.g.
         'Advantage_Fleet_LLC_-_May2026_-_Invoice_-_20260417.pdf'
         → 'Advantage Fleet LLC').
    """
    # Strategy 1: Redirect Health "Attn Company Broker" 3-column layout
    m = re.search(
        r'Attn\s+Company\s+Broker\s*\n\s*\S+\s+(.+?)\s{2,}',
        text, re.IGNORECASE
    )
    if m:
        candidate = m.group(1).strip()
        if candidate and len(candidate) > 1:
            return candidate

    # Strategy 2: explicit labels
    for pat in (
        r'Group\s*Name\s*:?\s*([^\n\r]{2,80})',
        r'(?:Bill|Sold)\s*To\s*:?\s*([^\n\r]{2,80})',
        r'Company\s*Name\s*:?\s*([^\n\r]{2,80})',
    ):
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

    # Strategy 3: filename
    if filename:
        base = os.path.basename(filename)
        base = os.path.splitext(base)[0]
        # Take everything before the first " - " or "_-_" separator
        for sep in (' - ', '_-_'):
            if sep in base:
                base = base.split(sep)[0]
                break
        # Convert underscores to spaces
        base = base.replace('_', ' ').strip()
        return base or None

    return None


def _find_invoice_amount(text: str) -> Optional[float]:
    """Find Total Invoice Amount / INVOICE AMOUNT."""
    for pat in (
        r'Total\s+Invoice\s+Amount\s*:?\s*\$?\s*\(?(-?[\d,]+\.?\d*)\)?',
        r'INVOICE\s+AMOUNT\s*:?\s*\$?\s*\(?(-?[\d,]+\.?\d*)\)?',
        r'Invoice\s+Total\s*:?\s*\$?\s*\(?(-?[\d,]+\.?\d*)\)?',
    ):
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return _to_number(m.group(1))
    return None

def _parse_cobra_members(text: str) -> set:
    cobra_names = set()
    # Find COBRA that appears as its own word on a line (the subsection header)
    # Must NOT be preceded by ( or word characters on the same line
    m = re.search(r'(?:^|\n)\s*COBRA\s*(?:\n|(?=\s+[A-Z]))', text)
    if not m:
        return cobra_names
    after = text[m.end():m.end() + 1000]
    chunk_m = re.search(r'Total\s*:', after, re.IGNORECASE)
    chunk = after[:chunk_m.start()] if chunk_m else after
    matches = re.findall(
        r'([A-Z][a-zA-Z]+)\s+([A-Z][a-zA-Z]+)\s+(?:PM|PC|PS|PF|EE)\b',
        chunk
    )
    for first, last in matches:
        nk = re.sub(r'[^a-z]', '', (first + last).lower())
        cobra_names.add(nk)
    return cobra_names

def _find_current_period_amount(text: str) -> Optional[float]:
    for pat in (
        r'Current\s+Period\s+Amount\s*:?\s*\$?\s*\(?(-?[\d,]+\.?\d*)\)?',
        r'Total\s+Current\s+Period\s+Amount\s*:?\s*\$?\s*\(?(-?[\d,]+\.?\d*)\)?',
    ):
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return _to_number(m.group(1))
    return None


def _find_adj_total(text: str, kind: str) -> Optional[float]:
    """kind = 'Charges' or 'Credits'. Handles parenthesized credits."""
    # Pattern allows either ( ... ) wrapping or - prefix
    cost = r'\(?\s*\$?\s*(-?[\d,]+\.?\d*)\s*\)?'
    for pat in (
        rf'Total\s+Adjustments\s*[-\u2013]\s*{kind}\s*:?\s*{cost}',
        rf'TOTAL\s+ADJUSTMENTS\s*[-\u2013]\s*{kind.upper()}\s*:?\s*{cost}',
        rf'Adjustments?\s*[-\u2013]\s*{kind}\s+Total\s*:?\s*{cost}',
    ):
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = _to_number(m.group(1))
            if val is None:
                continue
            # Check if the matched text had parentheses → negative
            matched_text = m.group(0)
            if '(' in matched_text and ')' in matched_text:
                val = -abs(val)
            elif kind == 'Credits' and val > 0:
                val = -abs(val)
            return val
    return None


def _to_number(s: str) -> Optional[float]:
    if s is None:
        return None
    try:
        return float(str(s).replace(',', '').replace('$', '').strip())
    except (ValueError, TypeError):
        return None


# ─── ROSTER SECTION PARSING ─────────────────────────────────────────
# Skip these lines: column headers, dividers, totals
_SKIP_PATTERNS = [
    re.compile(r'^\s*$'),                                   # blank
    re.compile(r'^\s*FIRST\s+NAME\b', re.IGNORECASE),       # column header
    re.compile(r'^\s*Division\s*:', re.IGNORECASE),         # division header
    re.compile(r'^\s*Total\s*:', re.IGNORECASE),            # subtotal
    re.compile(r'^\s*Subtotal\b', re.IGNORECASE),
    re.compile(r'^\s*Grand\s+Total\b', re.IGNORECASE),
    re.compile(r'^\s*INVOICE\s*#', re.IGNORECASE),
    re.compile(r'^\s*INVOICE\s+DATE', re.IGNORECASE),
    re.compile(r'^\s*GROUP\s*ID', re.IGNORECASE),
    re.compile(r'^\s*COVERAGE\s+PERIOD', re.IGNORECASE),
    re.compile(r'^\s*DATE\s*:', re.IGNORECASE),
    re.compile(r'^\s*DESCRIPTION\s', re.IGNORECASE),        # other table headers
]

# Sentinel patterns: when we hit one of these, we've left the roster section
_SECTION_BOUNDARIES = [
    re.compile(r'^\s*CURRENT\s+PERIOD\s*-', re.IGNORECASE),
    re.compile(r'^\s*ADJUSTMENTS?\s*-', re.IGNORECASE),
    re.compile(r'^\s*Redirect\s+Health', re.IGNORECASE),
    re.compile(r'^\s*Page\s+\d+\s+of', re.IGNORECASE),
]


def _parse_roster(text: str, section_header: str, has_coverage_month: bool) -> list:
    """
    Find a section by its header text, parse every member row beneath it.
    Stops at the next section header or end of text.

    Row formats:
      No coverage month:  FirstName  LastName  Tier  PlanCode  $Cost
      With coverage:      FirstName  LastName  Tier  PlanCode  CoverageMonth  $Cost
    """
    # Find section start
    pattern = re.compile(re.escape(section_header), re.IGNORECASE)
    m = pattern.search(text)
    if not m:
        return []

    # Take text from section header to end (will stop at next section ourselves)
    section_text = text[m.end():]
    lines = section_text.split('\n')

    rows = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # Stop at next section
        if any(p.match(stripped) for p in _SECTION_BOUNDARIES):
            break

        # Skip non-data lines
        if any(p.match(stripped) for p in _SKIP_PATTERNS):
            continue

        parsed = _parse_roster_row(stripped, has_coverage_month)
        if parsed:
            rows.append(parsed)

    return rows


def _parse_roster_row(line: str, has_coverage_month: bool) -> Optional[dict]:
    """
    Parse one line. Returns None if it doesn't match the expected shape.

    Token order (last to first):
      [cost] [coverage_month?] [plan_code] [tier] [...name...]

    Cost token handles: $123.45, 123.45, ($123.45), (123.45), -123.45
    Parenthesized = credit (negative).
    """
    # Cost pattern — order matters: optional (, optional $, digits, optional )
    cost_pattern = r'\(?\s*\$?\s*(-?[\d,]+\.\d{2})\s*\)?'

    if has_coverage_month:
        # Coverage month looks like "Mar-2026", "January 2026", "Mar 2026"
        # Pattern: name(greedy)  tier  plan  coverage_month  cost
        m = re.match(
            rf'^(.+?)\s+(\S+)\s+(\S+)\s+([A-Za-z]+[-\s]\d{{2,4}})\s+{cost_pattern}\s*$',
            line
        )
        if not m:
            return None
        name_part, tier, plan, cov_month, cost_str = m.groups()
    else:
        # Pattern: name(greedy)  tier  plan  cost
        m = re.match(
            rf'^(.+?)\s+(\S+)\s+(\S+)\s+{cost_pattern}\s*$',
            line
        )
        if not m:
            return None
        name_part, tier, plan, cost_str = m.groups()
        cov_month = None

    # Split name into first/last
    first, last = _split_name(name_part)
    if not first and not last:
        return None

    # Parse cost
    cost_clean = cost_str.replace(',', '').strip()
    try:
        cost = float(cost_clean)
    except ValueError:
        return None

    # Detect credit via parentheses or leading minus in the original line tail
    # Find the part of the line that contains the cost, check for ( ... )
    cost_idx = line.rfind(cost_str)
    if cost_idx > 0:
        tail = line[cost_idx - 2:cost_idx + len(cost_str) + 2]
        if '(' in tail and ')' in tail:
            cost = -abs(cost)

    entry = {
        'firstName': first,
        'lastName':  last,
        'tier':      tier,
        'planCode':  plan,
        'cost':      cost,
    }
    if cov_month is not None:
        entry['coverageMonth'] = cov_month
    return entry


def _split_name(name_part: str) -> tuple:
    """Split a name string into (firstName, lastName)."""
    name_part = name_part.strip()
    if not name_part:
        return ('', '')

    # "Last, First M" form
    if ',' in name_part:
        parts = [p.strip() for p in name_part.split(',', 1)]
        last = parts[0]
        first = parts[1] if len(parts) > 1 else ''
        return (first, last)

    # "First Last" or "First Middle Last" — split on first whitespace,
    # last word is last name, everything before is first (incl. middle).
    tokens = name_part.split()
    if len(tokens) == 1:
        return (tokens[0], '')
    if len(tokens) == 2:
        return (tokens[0], tokens[1])
    # 3+ tokens: First (Middle) Last — put middle into first to keep last clean
    return (' '.join(tokens[:-1]), tokens[-1])
