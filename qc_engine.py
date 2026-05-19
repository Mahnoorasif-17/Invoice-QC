"""
QC Engine — Excel-driven rules + PDF-driven adjustment validation.

Two responsibilities:
  1. run_qc()                  — Excel-driven rules (6 rules + late-adjustment scan).
  2. validate_pdf_adjustments() — PDF-driven validation: compare every charge/credit
                                  in the PDF against Excel. Each gets a status:
                                  ok / flag / incorrect.

Changes vs first version:
  • Rule 5 no longer fires when the start date is in the invoice month itself
    (regular billing already covers it).
  • Rule 6 no longer fires when the end date is in the invoice month itself.
  • New validation pass walks every PDF adjustment and validates it against Excel.
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


def month_floor(d: datetime) -> datetime:
    """Return the 1st of the given date's month."""
    return datetime(d.year, d.month, 1)


# ── CONTEXT ─────────────────────────────────────────────────────────
def auto_detect_month() -> tuple:
    today = datetime.now()
    day, m, y = today.day, today.month - 1, today.year
    if day >= 18:
        return ((m + 1) % 12, y + 1 if m == 11 else y)
    return (m, y)


def build_context(invoice_month: int, invoice_year: int) -> dict:
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
        'windowStart':  datetime(wSY, wSM + 1, 17),
        'windowEnd':    datetime(wEY, wEM + 1, 16, 23, 59, 59),
        'boundary60':   datetime(b6Y, b6M + 1, 1),
        'label':         f"{MF[iM]} {iY}",
        'prepLabel':     f"{MN[pM]} 18 – {MN[iM]} 17, {iY}",
        'windowLabel':   f"{MN[wSM]} 17 – {MN[wEM]} 16, {wEY}",
        'freeLabel':     f"{MN[b6M]} & {MN[wEM]}",
        'approvalLabel': f"Before {MN[b6M]} 1, {b6Y}",
    }


# ── QC ENGINE — EXCEL-DRIVEN RULES ──────────────────────────────────
def run_qc(rows: list, pdf_data: dict, ctx: dict) -> dict:
    """Run all 6 QC rules + late-adjustment scan + PDF-adjustment validation."""

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

        # Rule 1: Terminated but billed
        if ch and ch > 0 and st == 'Terminated':
            ed_s = ed.strftime('%m/%d/%Y') if ed else 'N/A'
            on_s = on.strftime('%m/%d/%Y') if on else 'N/A'
            issues.append({'sev': 'error',
                'msg': f"Terminated but billed {format_amount(ch)}. "
                       f"EndDate: {ed_s} | Entered EN: {on_s}"})

        # ── Rule 2 (FIXED): Active but not billed ──
        # Compare at MONTH level — someone starting May 10 should still be
        # billed in the May invoice, just like someone starting May 1.
        if (not ch or ch == 0) and st == 'Active' and sd \
                and month_floor(sd) <= ctx['invoiceStart'] \
                and pE and pE != '0':
            if not ed or month_floor(ed) >= ctx['invoiceStart']:
                issues.append({'sev': 'error',
                    'msg': f"Active since {sd.strftime('%m/%d/%Y')} but NOT billed in current period"})

        # ── Rule 3 (FIXED): Future start but billed ──
        # Only flag when the start MONTH is after the invoice month.
        # A May 10 start in a May invoice is fine — regular billing covers it.
        if ch and ch > 0 and sd and month_floor(sd) > ctx['invoiceStart']:
            issues.append({'sev': 'error',
                'msg': f"Future start {sd.strftime('%m/%d/%Y')} but billed {format_amount(ch)}"})

        # Rule 4: SKU mismatch / PEPM
        if ch and ch > 0 and pE and pS and pE != '0' and pS != '0':
            if pE != pS:
                issues.append({'sev': 'error',
                    'msg': f"SKU mismatch — EN: {pE} vs Invoice: {pS}"})
            elif ce is not None and abs(ce - ch) > 0.5:
                issues.append({'sev': 'pepm',
                    'msg': (f"Same SKU ({pE}), amount differs — EN: {format_amount(ce)} vs "
                            f"Invoice: {format_amount(ch)} (Δ{format_amount(abs(ce - ch))}) "
                            f"— possible PEPM")})

        # ── Rule 5 (FIXED) ──────────────────────────────────────────
        # Enrolled in data window, no adjustment charge — flag ONLY when the
        # start-date month is BEFORE the invoice month. If start is IN the
        # invoice month, regular billing already covers it (no adjustment needed).
        if eo and in_window(eo, ctx['windowStart'], ctx['windowEnd']) \
                and st == 'Active' and nk not in adj_charges:
            am = month_floor(sd) if sd else None
            if am is not None and am < ctx['invoiceStart']:
                needs_appr = am < ctx['boundary60']
                sev = 'approval' if needs_appr else 'warning'
                suffix = ' — NEEDS APPROVAL' if needs_appr else ''
                issues.append({'sev': sev,
                    'msg': (f"Enrolled {eo.strftime('%m/%d/%Y')} (data window). "
                            f"Adjustment for {format_month(am)} missing{suffix}")})

        # ── Rule 6 (FIXED) ──────────────────────────────────────────
        # Ended in data window, no adjustment credit — flag ONLY when the
        # end-date month is BEFORE the invoice month. If end is IN the
        # invoice month, regular billing already accounts for the partial month.
        if on and in_window(on, ctx['windowStart'], ctx['windowEnd']) \
                and nk not in adj_credits:
            cm = month_floor(ed) if ed else None
            if cm is not None and cm < ctx['invoiceStart']:
                needs_appr = cm < ctx['boundary60']
                sev = 'approval' if needs_appr else 'warning'
                suffix = ' — NEEDS APPROVAL' if needs_appr else ''
                issues.append({'sev': sev,
                    'msg': (f"Coverage ended {on.strftime('%m/%d/%Y')} (data window). "
                            f"Credit for {format_month(cm)} missing{suffix}")})

        if issues:
            member_issues.append({'name': name or '(no name)', 'iss': issues})

    # Late adjustments (older than 60-day boundary)
    late_adj = []
    for r in (pdf_data.get('adjustmentChargeRoster') or []):
        d = parse_month_string(r.get('coverageMonth'))
        if d and d < ctx['boundary60']:
            late_adj.append({
                'name': f"{r.get('firstName', '')} {r.get('lastName', '')}".strip(),
                'type': 'Charge', 'month': r.get('coverageMonth', ''),
                'cost': to_float(r.get('cost')) or 0, 'plan': r.get('planCode', '')})
    for r in (pdf_data.get('adjustmentCreditRoster') or []):
        d = parse_month_string(r.get('coverageMonth'))
        if d and d < ctx['boundary60']:
            late_adj.append({
                'name': f"{r.get('firstName', '')} {r.get('lastName', '')}".strip(),
                'type': 'Credit', 'month': r.get('coverageMonth', ''),
                'cost': to_float(r.get('cost')) or 0, 'plan': r.get('planCode', '')})

    # NEW: validate every PDF adjustment against Excel
    validations = validate_pdf_adjustments(rows, pdf_data, ctx)

    has_errors = any(any(i['sev'] == 'error' for i in m['iss']) for m in member_issues)
    has_warnings = any(any(i['sev'] in ('warning', 'pepm', 'approval') for i in m['iss'])
                       for m in member_issues)

    # Roll validation outcomes into the overall status
    if any(v['status'] == 'incorrect' for v in validations):
        has_errors = True
    if any(v['status'] == 'flag' for v in validations):
        has_warnings = True

    return {
        'groupId': pdf_data.get('groupId', '') or '',
        'groupName': pdf_data.get('groupName', '') or '',
        'invoiceAmount': pdf_data.get('invoiceAmount', 0) or 0,
        'currentPeriodAmount': pdf_data.get('currentPeriodAmount', 0) or 0,
        'adjCharges': pdf_data.get('adjustmentChargesTotal', 0) or 0,
        'adjCredits': pdf_data.get('adjustmentCreditsTotal', 0) or 0,
        'mi': member_issues,
        'la': late_adj,
        'validations': validations,
        'hasErrors': has_errors,
        'hasWarnings': not has_errors and (has_warnings or len(late_adj) > 0),
        'isClean': not has_errors and not has_warnings and len(late_adj) == 0,
    }


# ── PDF-DRIVEN VALIDATION ───────────────────────────────────────────
def validate_pdf_adjustments(rows: list, pdf_data: dict, ctx: dict) -> list:
    """
    For every PDF adjustment, look the member up in Excel and decide:
        ok        → adjustment is consistent with the Excel record
        flag      → suspicious / missing data / partial match — needs eyes
        incorrect → contradicts the Excel record outright

    Status feeds into the group's overall hasErrors / hasWarnings flag.
    """
    # Build Excel lookup by normalized name
    name_lookup = {}
    for row in rows:
        fn = str(row.get('FirstName') or '').strip()
        ln = str(row.get('LastName') or '').strip()
        nk = normalize_name(fn + ln)
        if nk:
            name_lookup[nk] = row

    def _validate(adj: dict, kind: str) -> dict:
        fn = str(adj.get('firstName', '')).strip()
        ln = str(adj.get('lastName', '')).strip()
        nk = normalize_name(fn + ln)
        full_name = f"{fn} {ln}".strip() or '(no name)'
        cov_str   = str(adj.get('coverageMonth', '') or '')
        cov_month = parse_month_string(cov_str)
        cost      = to_float(adj.get('cost')) or 0

        entry = {
            'name':     full_name,
            'type':     kind,
            'month':    cov_str or '?',
            'cost':     cost,
            'planCode': str(adj.get('planCode', '') or ''),
        }

        if nk not in name_lookup:
            entry['status'] = 'incorrect'
            entry['reason'] = 'Member not found in Excel roster'
            return entry

        if not cov_month:
            entry['status'] = 'flag'
            entry['reason'] = f'Could not parse coverage month: "{cov_str}"'
            return entry

        emp = name_lookup[nk]
        sd = parse_date(emp.get('StartDate'))
        ed = parse_date(emp.get('EndDate'))

        if not sd:
            entry['status'] = 'flag'
            entry['reason'] = 'No start date in Excel — cannot verify'
            return entry

        sd_month = month_floor(sd)

        # Before they started → wrong
        if cov_month < sd_month:
            entry['status'] = 'incorrect'
            entry['reason'] = (f'Coverage month {format_month(cov_month)} is before '
                               f'start date {sd.strftime("%m/%d/%Y")}')
            return entry

        # After they ended → wrong
        if ed:
            ed_month = month_floor(ed)
            if cov_month > ed_month:
                entry['status'] = 'incorrect'
                entry['reason'] = (f'Coverage month {format_month(cov_month)} is after '
                                   f'end date {ed.strftime("%m/%d/%Y")}')
                return entry

        # Adjustments are supposed to back-bill, so a current/future month is suspicious
        if cov_month >= ctx['invoiceStart']:
            entry['status'] = 'flag'
            entry['reason'] = (f'Adjustment for current/future month '
                               f'{format_month(cov_month)} — adjustments usually back-bill')
            return entry

        if kind == 'Credit':
            if not ed:
                entry['status'] = 'flag'
                entry['reason'] = 'Credit issued but member has no end date in Excel'
                return entry
            ed_month = month_floor(ed)
            if cov_month < ed_month:
                entry['status'] = 'flag'
                entry['reason'] = (f'Credit for {format_month(cov_month)} but coverage '
                                   f'didn\'t end until {ed.strftime("%m/%d/%Y")}')
                return entry
            entry['status'] = 'ok'
            entry['reason'] = f'Coverage ended {ed.strftime("%m/%d/%Y")} — credit valid'
            return entry

        # kind == 'Charge'
        entry['status'] = 'ok'
        entry['reason'] = (f'Within employment period (started {sd.strftime("%m/%d/%Y")}'
                           + (f', ended {ed.strftime("%m/%d/%Y")}' if ed else '') + ')')
        return entry

    validations = []
    for a in (pdf_data.get('adjustmentChargeRoster') or []):
        validations.append(_validate(a, 'Charge'))
    for a in (pdf_data.get('adjustmentCreditRoster') or []):
        validations.append(_validate(a, 'Credit'))
    return validations
