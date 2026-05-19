"""
QC Engine — direct port of the runQC() logic from the original HTML tool.
Pure functions, no I/O. Takes Excel rows + parsed PDF dict, returns a result dict.
"""

import re
from datetime import datetime
from typing import Optional, Any
import pandas as pd

MN = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
      'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
MF = ['January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December']
MONTH_MAP = {'jan': 0, 'feb': 1, 'mar': 2, 'apr': 3, 'may': 4, 'jun': 5,
             'jul': 6, 'aug': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dec': 11}


# ── DATE / STRING UTILS ─────────────────────────────────────────────
def parse_date(v: Any) -> Optional[datetime]:
    """Robustly parse a date from Excel cells (Timestamp, str, NaN, etc.)."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    if not s or s.lower() in ('nan', 'none', '0', 'nat', '0.0'):
        return None
    try:
        d = pd.to_datetime(s, errors='coerce')
        if pd.isna(d):
            return None
        return d.to_pydatetime().replace(tzinfo=None)
    except Exception:
        return None


def in_window(d: Optional[datetime], start: datetime, end: datetime) -> bool:
    return d is not None and start <= d <= end


def normalize_name(s: Any) -> str:
    return re.sub(r'[^a-z]', '', str(s or '').lower())


def parse_month_string(s: Any) -> Optional[datetime]:
    """Parse strings like 'Jan 2025' or 'January, 2025' to a datetime (1st of month)."""
    m = re.match(r'.*?([a-zA-Z]+)[^a-zA-Z0-9]*(\d{4})', str(s or ''))
    if not m:
        return None
    mi = MONTH_MAP.get(m.group(1).lower()[:3])
    if mi is None:
        return None
    return datetime(int(m.group(2)), mi + 1, 1)


def format_amount(n: Any) -> str:
    if n is None:
        return '—'
    try:
        return f"${float(n):,.0f}"
    except (ValueError, TypeError):
        return '—'


def format_month(d: Optional[datetime]) -> str:
    return d.strftime('%b %Y') if d else '?'


def to_float(v: Any) -> Optional[float]:
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    try:
        s = str(v).replace('$', '').replace(',', '').strip()
        if not s or s.lower() in ('nan', 'none', ''):
            return None
        return float(s)
    except (ValueError, TypeError):
        return None


# ── CONTEXT (date windows for invoice month) ────────────────────────
def auto_detect_month() -> tuple:
    """If today >= 18th, auto-shift to next month. Returns (month_0indexed, year)."""
    today = datetime.now()
    day, m, y = today.day, today.month - 1, today.year  # m is 0-indexed
    if day >= 18:
        return ((m + 1) % 12, y + 1 if m == 11 else y)
    return (m, y)


def build_context(invoice_month: int, invoice_year: int) -> dict:
    """
    Build all date windows for a given invoice month (0-indexed: 0=Jan, 11=Dec).
    Mirrors buildCtx() from the original JS exactly.
    """
    iM, iY = invoice_month, invoice_year

    wEM = 11 if iM == 0 else iM - 1
    wEY = iY - 1 if iM == 0 else iY

    wSM = 11 if wEM == 0 else wEM - 1
    wSY = wEY - 1 if wEM == 0 else wEY

    if iM <= 1:
        b6M, b6Y = iM + 10, iY - 1
    else:
        b6M, b6Y = iM - 2, iY

    pM = 11 if iM == 0 else iM - 1
    pY = iY - 1 if iM == 0 else iY

    return {
        'invoiceStart': datetime(iY, iM + 1, 1),
        'windowStart': datetime(wSY, wSM + 1, 17),
        'windowEnd': datetime(wEY, wEM + 1, 16, 23, 59, 59),
        'boundary60': datetime(b6Y, b6M + 1, 1),
        'label': f"{MF[iM]} {iY}",
        'prepLabel': f"{MN[pM]} 18 – {MN[iM]} 17, {iY}",
        'windowLabel': f"{MN[wSM]} 17 – {MN[wEM]} 16, {wEY}",
        'freeLabel': f"{MN[b6M]} & {MN[wEM]}",
        'approvalLabel': f"Before {MN[b6M]} 1, {b6Y}",
    }


# ── QC ENGINE ───────────────────────────────────────────────────────
def run_qc(rows: list, pdf_data: dict, ctx: dict) -> dict:
    """Run all 6 QC rules + late-adjustment scan for one group's invoice."""

    # Build adjustment lookup by normalized name
    adj_charges, adj_credits = {}, {}
    for r in (pdf_data.get('adjustmentChargeRoster') or []):
        k = normalize_name(str(r.get('firstName', '')) + str(r.get('lastName', '')))
        adj_charges.setdefault(k, []).append(r)
    for r in (pdf_data.get('adjustmentCreditRoster') or []):
        k = normalize_name(str(r.get('firstName', '')) + str(r.get('lastName', '')))
        adj_credits.setdefault(k, []).append(r)

    member_issues = []
    for row in rows:
        fn = str(row.get('FirstName') or '').strip()
        ln = str(row.get('LastName') or '').strip()
        name = f"{fn} {ln}".strip()
        nk = normalize_name(fn + ln)

        ch = to_float(row.get('Charge'))
        ce = to_float(row.get('PlanCost'))
        pE = str(row.get('CarrierPlanCode') or '').strip()
        pS = str(row.get('PlanSKUCSV') or '').strip()
        st = str(row.get('EmploymentStatus') or '').strip()
        sd = parse_date(row.get('StartDate'))
        eo = parse_date(row.get('EnrolledOn'))
        ed = parse_date(row.get('EndDate'))
        on = parse_date(row.get('EndedOn'))

        issues = []

        # Rule 1: Terminated but still billed
        if ch and ch > 0 and st == 'Terminated':
            ed_s = ed.strftime('%m/%d/%Y') if ed else 'N/A'
            on_s = on.strftime('%m/%d/%Y') if on else 'N/A'
            issues.append({
                'sev': 'error',
                'msg': f"Terminated but billed {format_amount(ch)}. EndDate: {ed_s} | Entered EN: {on_s}"
            })

        # Rule 2: Active since start date but not billed
        if (not ch or ch == 0) and st == 'Active' and sd and sd <= ctx['invoiceStart'] \
                and pE and pE != '0':
            if not ed or ed >= ctx['invoiceStart']:
                issues.append({
                    'sev': 'error',
                    'msg': f"Active since {sd.strftime('%m/%d/%Y')} but NOT billed in current period"
                })

        # Rule 3: Future start date but already billed
        if ch and ch > 0 and sd and sd > ctx['invoiceStart']:
            issues.append({
                'sev': 'error',
                'msg': f"Future start {sd.strftime('%m/%d/%Y')} but billed {format_amount(ch)}"
            })

        # Rule 4: SKU mismatch OR same-SKU/different-amount (PEPM)
        if ch and ch > 0 and pE and pS and pE != '0' and pS != '0':
            if pE != pS:
                issues.append({
                    'sev': 'error',
                    'msg': f"SKU mismatch — EN: {pE} vs Invoice: {pS}"
                })
            elif ce is not None and abs(ce - ch) > 0.5:
                issues.append({
                    'sev': 'pepm',
                    'msg': (f"Same SKU ({pE}), amount differs — EN: {format_amount(ce)} vs "
                            f"Invoice: {format_amount(ch)} (Δ{format_amount(abs(ce - ch))}) "
                            f"— possible PEPM")
                })

        # Rule 5: Enrolled in data window, but no adjustment charge present
        if eo and in_window(eo, ctx['windowStart'], ctx['windowEnd']) \
                and st == 'Active' and nk not in adj_charges:
            am = datetime(sd.year, sd.month, 1) if sd else None
            needs_appr = am is not None and am < ctx['boundary60']
            sev = 'approval' if needs_appr else 'warning'
            suffix = ' — NEEDS APPROVAL' if needs_appr else ''
            issues.append({
                'sev': sev,
                'msg': (f"Enrolled {eo.strftime('%m/%d/%Y')} (data window). "
                        f"Adjustment for {format_month(am)} missing{suffix}")
            })

        # Rule 6: Ended in data window, but no adjustment credit present
        if on and in_window(on, ctx['windowStart'], ctx['windowEnd']) \
                and nk not in adj_credits:
            cm = datetime(ed.year, ed.month, 1) if ed else None
            needs_appr = cm is not None and cm < ctx['boundary60']
            sev = 'approval' if needs_appr else 'warning'
            suffix = ' — NEEDS APPROVAL' if needs_appr else ''
            issues.append({
                'sev': sev,
                'msg': (f"Coverage ended {on.strftime('%m/%d/%Y')} (data window). "
                        f"Credit for {format_month(cm)} missing{suffix}")
            })

        if issues:
            member_issues.append({'name': name or '(no name)', 'iss': issues})

    # Late adjustments — anything older than the 60-day boundary
    late_adj = []
    for r in (pdf_data.get('adjustmentChargeRoster') or []):
        d = parse_month_string(r.get('coverageMonth'))
        if d and d < ctx['boundary60']:
            late_adj.append({
                'name': f"{r.get('firstName', '')} {r.get('lastName', '')}".strip(),
                'type': 'Charge',
                'month': r.get('coverageMonth', ''),
                'cost': to_float(r.get('cost')) or 0,
                'plan': r.get('planCode', '')
            })
    for r in (pdf_data.get('adjustmentCreditRoster') or []):
        d = parse_month_string(r.get('coverageMonth'))
        if d and d < ctx['boundary60']:
            late_adj.append({
                'name': f"{r.get('firstName', '')} {r.get('lastName', '')}".strip(),
                'type': 'Credit',
                'month': r.get('coverageMonth', ''),
                'cost': to_float(r.get('cost')) or 0,
                'plan': r.get('planCode', '')
            })

    has_errors = any(any(i['sev'] == 'error' for i in m['iss']) for m in member_issues)
    has_warnings = any(any(i['sev'] in ('warning', 'pepm', 'approval') for i in m['iss'])
                       for m in member_issues)

    return {
        'groupId': pdf_data.get('groupId', '') or '',
        'groupName': pdf_data.get('groupName', '') or '',
        'invoiceAmount': pdf_data.get('invoiceAmount', 0) or 0,
        'currentPeriodAmount': pdf_data.get('currentPeriodAmount', 0) or 0,
        'adjCharges': pdf_data.get('adjustmentChargesTotal', 0) or 0,
        'adjCredits': pdf_data.get('adjustmentCreditsTotal', 0) or 0,
        'mi': member_issues,
        'la': late_adj,
        'hasErrors': has_errors,
        'hasWarnings': not has_errors and (has_warnings or len(late_adj) > 0),
        'isClean': not has_errors and not has_warnings and len(late_adj) == 0,
    }
