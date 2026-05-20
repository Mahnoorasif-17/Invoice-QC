"""
QC Engine — Excel is the source of truth.

Two separate concerns now:

  1. BILLING issues (run_qc → member_issues)
     Rules 1-4: terminated/billed, active/not-billed, future-start/billed,
     SKU mismatch / PEPM. These are Excel-internal + Excel-vs-PDF-billing checks.

  2. ADJUSTMENT/CREDIT validation (validate_adjustments → validations)
     Unified Excel→PDF report. Three statuses:
       - ok        : Excel predicted this adjustment AND PDF has it
       - flag      : Excel predicted this adjustment but PDF is MISSING it,
                     OR a PDF entry is suspicious (can't parse month, etc.)
       - incorrect : PDF has an adjustment that contradicts Excel
                     (wrong member, before start date, after end date, etc.)

There is NO duplication — Rules 5 and 6 from the old version (which lived in
member_issues) are gone. All adjustment logic is now in validate_adjustments().
"""

import calendar
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


# ── UTILS ───────────────────────────────────────────────────────────
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
    return datetime(d.year, d.month, 1)


def is_last_day_of_month(d: datetime) -> bool:
    """True if d is the last calendar day of its month (e.g. Apr 30, Feb 28/29)."""
    return d.day == calendar.monthrange(d.year, d.month)[1]


def next_month(d: datetime) -> datetime:
    """Return the 1st of the month after d."""
    if d.month == 12:
        return datetime(d.year + 1, 1, 1)
    return datetime(d.year, d.month + 1, 1)


def months_between(start_month: datetime, end_exclusive: datetime):
    """Yield month-1st datetimes from start_month up to but not including end_exclusive."""
    cur = start_month
    while cur < end_exclusive:
        yield cur
        cur = next_month(cur)


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


# ────────────────────────────────────────────────────────────────────
# MAIN QC — BILLING RULES (1-4) ONLY
# Adjustment/credit handling is delegated to validate_adjustments below.
# ────────────────────────────────────────────────────────────────────
def run_qc(rows: list, pdf_data: dict, ctx: dict) -> dict:
    member_issues = []
    for row in rows:
        fn = str(row.get('FirstName') or '').strip()
        ln = str(row.get('LastName') or '').strip()
        name = f"{fn} {ln}".strip()

        ch = to_float(row.get('Charge'))
        ce = to_float(row.get('PlanCost'))
        pE = str(row.get('CarrierPlanCode') or '').strip()
        pS = str(row.get('PlanSKUCSV') or '').strip()
        st = str(row.get('EmploymentStatus') or '').strip()
        sd = parse_date(row.get('StartDate'))
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

        # Rule 2: Active since (month ≤ invoice month) but not billed
        if (not ch or ch == 0) and st == 'Active' and sd \
                and month_floor(sd) <= ctx['invoiceStart'] \
                and pE and pE != '0':
            if not ed or month_floor(ed) >= ctx['invoiceStart']:
                issues.append({'sev': 'error',
                    'msg': f"Active since {sd.strftime('%m/%d/%Y')} but NOT billed in current period"})

        # Rule 3: Future start (in a month AFTER invoice month) but billed
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

        if issues:
            member_issues.append({'name': name or '(no name)', 'iss': issues})

    # Late adjustments (older than 60-day boundary) — for the approval section
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

    # Adjustment/credit validation (the new unified report)
    validations = validate_adjustments(rows, pdf_data, ctx)

    has_errors = any(any(i['sev'] == 'error' for i in m['iss']) for m in member_issues)
    has_warnings = any(any(i['sev'] in ('warning', 'pepm', 'approval') for i in m['iss'])
                       for m in member_issues)

    # Roll validation outcomes into overall status
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


# ────────────────────────────────────────────────────────────────────
# ADJUSTMENT/CREDIT VALIDATION — EXCEL → PDF, UNIFIED REPORT
# ────────────────────────────────────────────────────────────────────
def validate_adjustments(rows: list, pdf_data: dict, ctx: dict) -> list:
    """
    Returns a list of validation entries, each with:
        status   : 'ok' | 'flag' | 'incorrect'
        type     : 'Charge' | 'Credit'
        name     : member name
        month    : coverage month string
        cost     : amount
        plan     : plan code
        reason   : human-readable explanation

    Two passes:
      Pass A (Excel → PDF): For each member, derive expected adjustments from
                            their Excel dates. Match against PDF entries.
                            Match  = OK.   Missing = FLAG.
      Pass B (PDF → Excel): For each PDF entry not matched in Pass A, validate
                            against Excel. Unmatched but valid = FLAG.
                            Contradicts Excel = INCORRECT.
    """
    # Index PDF entries by normalized name
    pdf_charges_by_name = {}  # nk -> list[(month_dt|None, raw_row)]
    pdf_credits_by_name = {}
    for r in (pdf_data.get('adjustmentChargeRoster') or []):
        nk = normalize_name(str(r.get('firstName', '')) + str(r.get('lastName', '')))
        pdf_charges_by_name.setdefault(nk, []).append(
            (parse_month_string(r.get('coverageMonth')), r))
    for r in (pdf_data.get('adjustmentCreditRoster') or []):
        nk = normalize_name(str(r.get('firstName', '')) + str(r.get('lastName', '')))
        pdf_credits_by_name.setdefault(nk, []).append(
            (parse_month_string(r.get('coverageMonth')), r))

    # Track which PDF entries were matched to an Excel expectation
    matched_charges, matched_credits = set(), set()  # set of id() of raw_row

    # Excel lookup
    name_to_row = {}
    for row in rows:
        fn = str(row.get('FirstName') or '').strip()
        ln = str(row.get('LastName') or '').strip()
        nk = normalize_name(fn + ln)
        if nk:
            name_to_row[nk] = row

    validations = []

    # ── PASS A: Excel-driven expectations ───────────────────────────
    for row in rows:
        fn = str(row.get('FirstName') or '').strip()
        ln = str(row.get('LastName') or '').strip()
        name = f"{fn} {ln}".strip() or '(no name)'
        nk = normalize_name(fn + ln)

        st = str(row.get('EmploymentStatus') or '').strip()
        sd = parse_date(row.get('StartDate'))
        eo = parse_date(row.get('EnrolledOn'))
        ed = parse_date(row.get('EndDate'))
        on = parse_date(row.get('EndedOn'))
        plan = str(row.get('CarrierPlanCode') or '').strip()
        plan_cost = to_float(row.get('PlanCost')) or 0

        # ── Expected adjustment CHARGES: enrolled in data window AND start
        #    month is BEFORE the invoice month. Predict ONE charge per missed
        #    month — from start_month through (invoice_month − 1) inclusive.
        if eo and in_window(eo, ctx['windowStart'], ctx['windowEnd']) \
                and st == 'Active' and sd:
            sd_month = month_floor(sd)
            if sd_month < ctx['invoiceStart']:
                for predicted_m in months_between(sd_month, ctx['invoiceStart']):
                    # Find unmatched PDF charge for this exact month
                    found_row = None
                    for m, r in pdf_charges_by_name.get(nk, []):
                        if m == predicted_m and id(r) not in matched_charges:
                            found_row = r
                            matched_charges.add(id(r))
                            break

                    if found_row is not None:
                        validations.append({
                            'status': 'ok',
                            'type':   'Charge',
                            'name':   name,
                            'month':  format_month(predicted_m),
                            'cost':   to_float(found_row.get('cost')) or 0,
                            'plan':   str(found_row.get('planCode', '') or ''),
                            'reason': (f"Expected (enrolled {eo.strftime('%m/%d/%Y')}, "
                                       f"start {sd.strftime('%m/%d/%Y')}) — found in PDF"),
                        })
                    else:
                        needs_appr = predicted_m < ctx['boundary60']
                        suffix = ' — NEEDS APPROVAL (>60 days)' if needs_appr else ''
                        validations.append({
                            'status': 'flag',
                            'type':   'Charge',
                            'name':   name,
                            'month':  format_month(predicted_m),
                            'cost':   plan_cost,
                            'plan':   plan,
                            'reason': (f"MISSING — enrolled {eo.strftime('%m/%d/%Y')} (data window), "
                                       f"start {sd.strftime('%m/%d/%Y')} → expected charge for "
                                       f"{format_month(predicted_m)} but PDF has none{suffix}"),
                        })

        # ── Expected CREDIT: ended in data window AND end month is BEFORE
        #    the invoice month (in-month ends are covered by regular billing).
        #    ALSO skip if end date is the LAST day of its month — the full
        #    month was earned, so no credit is owed.
        if on and in_window(on, ctx['windowStart'], ctx['windowEnd']) and ed:
            ed_month = month_floor(ed)
            if ed_month < ctx['invoiceStart'] and not is_last_day_of_month(ed):
                found_row = None
                for m, r in pdf_credits_by_name.get(nk, []):
                    if m == ed_month and id(r) not in matched_credits:
                        found_row = r
                        matched_credits.add(id(r))
                        break

                if found_row is not None:
                    validations.append({
                        'status': 'ok',
                        'type':   'Credit',
                        'name':   name,
                        'month':  format_month(ed_month),
                        'cost':   to_float(found_row.get('cost')) or 0,
                        'plan':   str(found_row.get('planCode', '') or ''),
                        'reason': (f"Expected (ended {ed.strftime('%m/%d/%Y')}, "
                                   f"entered EN {on.strftime('%m/%d/%Y')}) — found in PDF"),
                    })
                else:
                    needs_appr = ed_month < ctx['boundary60']
                    suffix = ' — NEEDS APPROVAL (>60 days)' if needs_appr else ''
                    validations.append({
                        'status': 'flag',
                        'type':   'Credit',
                        'name':   name,
                        'month':  format_month(ed_month),
                        'cost':   plan_cost,
                        'plan':   plan,
                        'reason': (f"MISSING — ended {ed.strftime('%m/%d/%Y')}, entered EN "
                                   f"{on.strftime('%m/%d/%Y')} (data window) → expected credit "
                                   f"for {format_month(ed_month)} but PDF has none{suffix}"),
                    })

    # ── PASS B: PDF entries not matched in Pass A ───────────────────
    def _classify_pdf_entry(nk: str, m: Optional[datetime], r: dict, kind: str) -> dict:
        full_name = f"{r.get('firstName', '')} {r.get('lastName', '')}".strip() or '(no name)'
        cov_str = str(r.get('coverageMonth', '') or '')
        cost = to_float(r.get('cost')) or 0
        plan = str(r.get('planCode', '') or '')

        base = {'type': kind, 'name': full_name,
                'month': cov_str or '?', 'cost': cost, 'plan': plan}

        # Member not in Excel → INCORRECT
        if nk not in name_to_row:
            return {**base, 'status': 'incorrect',
                    'reason': f'Member not found in Excel roster — PDF has {kind.lower()} '
                              f'but no record exists'}

        # Couldn't parse coverage month → FLAG
        if m is None:
            return {**base, 'status': 'flag',
                    'reason': f'Could not parse coverage month: "{cov_str}"'}

        emp = name_to_row[nk]
        sd = parse_date(emp.get('StartDate'))
        ed = parse_date(emp.get('EndDate'))

        if sd is None:
            return {**base, 'status': 'flag',
                    'reason': 'No start date in Excel — cannot verify'}

        sd_month = month_floor(sd)

        # Coverage before start → INCORRECT
        if m < sd_month:
            return {**base, 'status': 'incorrect',
                    'reason': (f'Coverage month {format_month(m)} is before start date '
                               f'{sd.strftime("%m/%d/%Y")}')}

        # Coverage after end → INCORRECT
        if ed is not None and m > month_floor(ed):
            return {**base, 'status': 'incorrect',
                    'reason': (f'Coverage month {format_month(m)} is after end date '
                               f'{ed.strftime("%m/%d/%Y")}')}

        # For credits: member should have an end date
        if kind == 'Credit' and ed is None:
            return {**base, 'status': 'flag',
                    'reason': 'Credit in PDF but member has no end date in Excel'}

        # Within employment period but Excel didn't predict it — FLAG with note
        return {**base, 'status': 'flag',
                'reason': (f'Within employment period but not predicted by data window — '
                           f'review (started {sd.strftime("%m/%d/%Y")}'
                           + (f', ended {ed.strftime("%m/%d/%Y")}' if ed else '')
                           + ')')}

    for nk, entries in pdf_charges_by_name.items():
        for m, r in entries:
            if id(r) in matched_charges:
                continue
            validations.append(_classify_pdf_entry(nk, m, r, 'Charge'))

    for nk, entries in pdf_credits_by_name.items():
        for m, r in entries:
            if id(r) in matched_credits:
                continue
            validations.append(_classify_pdf_entry(nk, m, r, 'Credit'))

    # Sort: incorrect first, then flag, then ok
    sort_order = {'incorrect': 0, 'flag': 1, 'ok': 2}
    validations.sort(key=lambda v: (sort_order[v['status']], v['type'], v['name']))

    return validations
