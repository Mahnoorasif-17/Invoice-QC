"""
QC Engine — Excel is the source of truth.

BILLING RULES (1-4): walk Excel rows, flag billing problems.
ADJUSTMENT / CREDIT VALIDATION: per-member, with these rules:

  ── Did anything fall in the data window? ───
  A member is "in scope" for adjustment validation if EITHER:
    • EnrolledOn is in the data window  → expect charges for past months they started
    • EndedOn    is in the data window  → expect credits for months billed AFTER end

  ── Credits ───
  • EndDate in any past month → NO credit needed for that end month
    (member used some insurance that month → full month billed → no credit owed)
  • But credit IS needed for any month billed AFTER the end month
    (e.g. EndDate Feb 28 → no credit for Feb, but credits needed for Mar, Apr if billed)

  ── Charges ───
  • StartDate in any past month + EnrolledOn in data window → charge needed for
    every month from start month through invoice_month−1 inclusive.

Per-member verdict (one row in the report):
  • OK            → all expected adjustments are in the PDF
  • NO_ADJ_NEEDED → in scope (data window) but no adjustments needed AND PDF has none
  • MISSING       → expected some adjustments, but PDF is missing one or more
  • INCORRECT     → PDF has entries that contradict Excel (wrong member, before
                    start, after end+1month, etc.)
"""

import calendar
import re
from datetime import datetime
from typing import Optional, Any, List, Dict
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


def next_month(d: datetime) -> datetime:
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
# MAIN QC — BILLING RULES (1-4)
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

    # Late adjustments (older than 60-day boundary) — listed separately
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

    # Adjustment/credit validation — per-member report
    validations = validate_adjustments_per_member(rows, pdf_data, ctx)

    has_errors = any(any(i['sev'] == 'error' for i in m['iss']) for m in member_issues)
    has_warnings = any(any(i['sev'] in ('warning', 'pepm', 'approval') for i in m['iss'])
                       for m in member_issues)

    # Roll validation outcomes into overall status
    if any(v['status'] == 'incorrect' for v in validations):
        has_errors = True
    if any(v['status'] == 'missing' for v in validations):
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
# ADJUSTMENT/CREDIT VALIDATION — PER MEMBER (one row each)
# ────────────────────────────────────────────────────────────────────
def validate_adjustments_per_member(rows: list, pdf_data: dict, ctx: dict) -> list:
    """
    Returns one validation entry per relevant member, each shaped like:
      {
        'status':    'ok' | 'missing' | 'incorrect' | 'no_adj_needed',
        'type':      'Charge' | 'Credit' | 'Charge & Credit' | 'None',
        'name':      'First Last',
        'months':    'Mar 2026, Apr 2026'   ← all expected months joined
        'expected':  ['Mar 2026', 'Apr 2026'],
        'found':     ['Mar 2026'],
        'missing':   ['Apr 2026'],
        'cost':      total amount expected (sum)
        'plan':      plan code from Excel
        'reason':    human-readable explanation
      }
    """
    # Index PDF entries by normalized name
    pdf_charges_by_name: Dict[str, list] = {}
    pdf_credits_by_name: Dict[str, list] = {}
    for r in (pdf_data.get('adjustmentChargeRoster') or []):
        nk = normalize_name(str(r.get('firstName', '')) + str(r.get('lastName', '')))
        pdf_charges_by_name.setdefault(nk, []).append(
            (parse_month_string(r.get('coverageMonth')), r))
    for r in (pdf_data.get('adjustmentCreditRoster') or []):
        nk = normalize_name(str(r.get('firstName', '')) + str(r.get('lastName', '')))
        pdf_credits_by_name.setdefault(nk, []).append(
            (parse_month_string(r.get('coverageMonth')), r))

    # Track matched PDF rows so we can detect unexpected leftovers
    matched_charge_ids = set()
    matched_credit_ids = set()

    # Excel lookup
    name_to_row: Dict[str, dict] = {}
    for row in rows:
        fn = str(row.get('FirstName') or '').strip()
        ln = str(row.get('LastName') or '').strip()
        nk = normalize_name(fn + ln)
        if nk:
            name_to_row[nk] = row

    validations: List[dict] = []

    # ─── PASS A: Walk every Excel member in scope ────────────────────
    for row in rows:
        fn = str(row.get('FirstName') or '').strip()
        ln = str(row.get('LastName') or '').strip()
        name = f"{fn} {ln}".strip() or '(no name)'
        nk = normalize_name(fn + ln)
        plan = str(row.get('CarrierPlanCode') or '').strip()
        plan_cost = to_float(row.get('PlanCost')) or 0

        sd = parse_date(row.get('StartDate'))
        eo = parse_date(row.get('EnrolledOn'))
        ed = parse_date(row.get('EndDate'))
        on = parse_date(row.get('EndedOn'))
        st = str(row.get('EmploymentStatus') or '').strip()

        eo_in_window = eo is not None and in_window(eo, ctx['windowStart'], ctx['windowEnd'])
        on_in_window = on is not None and in_window(on, ctx['windowStart'], ctx['windowEnd'])

        # Skip members who aren't relevant — neither enroll nor end fell in the data window
        if not eo_in_window and not on_in_window:
            continue

        # ── Build the list of expected CHARGES ────────────────────────
        expected_charge_months: List[datetime] = []
        if eo_in_window and st == 'Active' and sd:
            sd_month = month_floor(sd)
            if sd_month < ctx['invoiceStart']:
                # Charge needed for every month from start through invoice_month-1
                for m in months_between(sd_month, ctx['invoiceStart']):
                    expected_charge_months.append(m)

        # ── Build the list of expected CREDITS ────────────────────────
        # Rule: NO credit for the end month itself (any day = used insurance,
        # full month billed).  Credit IS needed for any month billed AFTER
        # the end month — i.e. (end_month + 1) through (invoice_month − 1).
        expected_credit_months: List[datetime] = []
        if on_in_window and ed:
            ed_month = month_floor(ed)
            first_credit_month = next_month(ed_month)
            if first_credit_month < ctx['invoiceStart']:
                for m in months_between(first_credit_month, ctx['invoiceStart']):
                    expected_credit_months.append(m)

        # ── Match PDF entries to expectations ─────────────────────────
        found_charges: List[datetime] = []
        missing_charges: List[datetime] = []
        for em in expected_charge_months:
            matched = None
            for m, r in pdf_charges_by_name.get(nk, []):
                if m == em and id(r) not in matched_charge_ids:
                    matched = r
                    matched_charge_ids.add(id(r))
                    break
            (found_charges if matched else missing_charges).append(em)

        found_credits: List[datetime] = []
        missing_credits: List[datetime] = []
        for em in expected_credit_months:
            matched = None
            for m, r in pdf_credits_by_name.get(nk, []):
                if m == em and id(r) not in matched_credit_ids:
                    matched = r
                    matched_credit_ids.add(id(r))
                    break
            (found_credits if matched else missing_credits).append(em)

        # ── Detect unexpected PDF entries for this member ─────────────
        # Any PDF charge/credit for this person that doesn't match an
        # expected month is suspicious. We'll classify it here too so the
        # per-member row owns its full verdict.
        unexpected_pdf: List[dict] = []
        for m, r in pdf_charges_by_name.get(nk, []):
            if id(r) in matched_charge_ids:
                continue
            unexpected_pdf.append({
                'kind': 'Charge', 'month_dt': m, 'row': r,
                'month_str': str(r.get('coverageMonth', '') or '?'),
                'cost': to_float(r.get('cost')) or 0,
            })
            matched_charge_ids.add(id(r))  # mark consumed so Pass B skips
        for m, r in pdf_credits_by_name.get(nk, []):
            if id(r) in matched_credit_ids:
                continue
            unexpected_pdf.append({
                'kind': 'Credit', 'month_dt': m, 'row': r,
                'month_str': str(r.get('coverageMonth', '') or '?'),
                'cost': to_float(r.get('cost')) or 0,
            })
            matched_credit_ids.add(id(r))

        # ── Decide the per-member verdict ─────────────────────────────
        has_expected = bool(expected_charge_months or expected_credit_months)
        has_missing = bool(missing_charges or missing_credits)
        has_unexpected = bool(unexpected_pdf)

        # Build label parts
        type_parts = []
        if expected_charge_months or any(u['kind'] == 'Charge' for u in unexpected_pdf):
            type_parts.append('Charge')
        if expected_credit_months or any(u['kind'] == 'Credit' for u in unexpected_pdf):
            type_parts.append('Credit')
        type_label = ' & '.join(type_parts) if type_parts else 'None'

        # Build month list
        all_months = (
            [format_month(m) for m in (expected_charge_months + expected_credit_months)]
            + [u['month_str'] for u in unexpected_pdf]
        )
        months_str = ', '.join(all_months) if all_months else '—'

        # ── Status decision tree ──────────────────────────────────────
        if has_unexpected:
            # Validate each unexpected entry against employment dates
            problem_reasons = []
            for u in unexpected_pdf:
                m = u['month_dt']
                if m is None:
                    problem_reasons.append(
                        f"Could not parse PDF coverage month {u['month_str']!r}")
                    continue
                if sd and m < month_floor(sd):
                    problem_reasons.append(
                        f"PDF has {u['kind']} for {format_month(m)} "
                        f"but member started {sd.strftime('%m/%d/%Y')}")
                elif ed and m > month_floor(ed) and u['kind'] == 'Charge':
                    problem_reasons.append(
                        f"PDF has {u['kind']} for {format_month(m)} "
                        f"but member ended {ed.strftime('%m/%d/%Y')}")
                elif u['kind'] == 'Credit' and not ed:
                    problem_reasons.append(
                        f"PDF has Credit for {format_month(m)} "
                        f"but member has no end date in Excel")
                else:
                    problem_reasons.append(
                        f"PDF has unexpected {u['kind']} for {format_month(m)}")
            status = 'incorrect'
            reason = '; '.join(problem_reasons)

        elif has_missing:
            parts = []
            if missing_charges:
                parts.append(f"Missing Charge(s) for {', '.join(format_month(m) for m in missing_charges)}")
            if missing_credits:
                parts.append(f"Missing Credit(s) for {', '.join(format_month(m) for m in missing_credits)}")
            # Note approval requirements
            approval_months = [m for m in (missing_charges + missing_credits) if m < ctx['boundary60']]
            if approval_months:
                parts.append(f"NEEDS APPROVAL — older than 60 days")
            status = 'missing'
            reason = '; '.join(parts)

        elif has_expected:
            # Everything expected was found
            parts = []
            if eo_in_window:
                parts.append(f"enrolled {eo.strftime('%m/%d/%Y')}")
            if on_in_window:
                parts.append(f"entered EN {on.strftime('%m/%d/%Y')}")
            if sd:
                parts.append(f"start {sd.strftime('%m/%d/%Y')}")
            if ed:
                parts.append(f"end {ed.strftime('%m/%d/%Y')}")
            status = 'ok'
            reason = f"All expected adjustments found in PDF ({', '.join(parts)})"

        else:
            # In data window, but nothing should have been adjusted
            # AND nothing unexpected in PDF → NO ADJ NEEDED
            reasons = []
            if eo_in_window and (not sd or month_floor(sd) >= ctx['invoiceStart']):
                if sd:
                    reasons.append(f"enrolled {eo.strftime('%m/%d/%Y')} but start "
                                   f"{sd.strftime('%m/%d/%Y')} is in invoice month — "
                                   f"regular billing covers")
                else:
                    reasons.append(f"enrolled {eo.strftime('%m/%d/%Y')} but no start date")
            if on_in_window and ed and month_floor(ed) < ctx['invoiceStart']:
                # End in past month, no month billed after end → no credit needed
                reasons.append(f"ended {ed.strftime('%m/%d/%Y')} — used insurance that month, "
                               f"no credit needed")
            if on_in_window and ed and month_floor(ed) >= ctx['invoiceStart']:
                reasons.append(f"ended {ed.strftime('%m/%d/%Y')} — in invoice month, "
                               f"regular billing covers")
            status = 'no_adj_needed'
            reason = '; '.join(reasons) if reasons else 'In data window but no adjustments needed'

        # ── Compute total expected amount and unmatched-entry details
        expected_amount = (
            (len(expected_charge_months) * plan_cost)
            + sum(abs(u['cost']) for u in unexpected_pdf if u['kind'] == 'Charge')
        )
        # For credits we keep signed
        validations.append({
            'status':   status,
            'type':     type_label,
            'name':     name,
            'months':   months_str,
            'expected_charges': [format_month(m) for m in expected_charge_months],
            'expected_credits': [format_month(m) for m in expected_credit_months],
            'found_charges':    [format_month(m) for m in found_charges],
            'found_credits':    [format_month(m) for m in found_credits],
            'missing_charges':  [format_month(m) for m in missing_charges],
            'missing_credits':  [format_month(m) for m in missing_credits],
            'unexpected_pdf':   [{'kind': u['kind'], 'month': u['month_str'],
                                  'cost': u['cost']} for u in unexpected_pdf],
            'cost':     expected_amount,
            'plan':     plan,
            'reason':   reason,
        })

    # ─── PASS B: PDF entries that didn't match any Excel member ──────
    # These are members in the PDF but not in Excel — INCORRECT
    pdf_only_members: Dict[str, List[dict]] = {}
    for nk, entries in pdf_charges_by_name.items():
        if nk not in name_to_row:
            for m, r in entries:
                pdf_only_members.setdefault(nk, []).append({
                    'kind': 'Charge', 'month_dt': m, 'row': r,
                    'month_str': str(r.get('coverageMonth', '') or '?'),
                    'cost': to_float(r.get('cost')) or 0,
                })
    for nk, entries in pdf_credits_by_name.items():
        if nk not in name_to_row:
            for m, r in entries:
                pdf_only_members.setdefault(nk, []).append({
                    'kind': 'Credit', 'month_dt': m, 'row': r,
                    'month_str': str(r.get('coverageMonth', '') or '?'),
                    'cost': to_float(r.get('cost')) or 0,
                })

    for nk, entries in pdf_only_members.items():
        # Reconstruct name from the PDF rows
        sample = entries[0]['row']
        name = f"{sample.get('firstName', '')} {sample.get('lastName', '')}".strip() or '(no name)'
        type_parts = sorted({e['kind'] for e in entries})
        type_label = ' & '.join(type_parts)
        months_str = ', '.join(e['month_str'] for e in entries)
        plan = str(sample.get('planCode', '') or '')

        validations.append({
            'status':   'incorrect',
            'type':     type_label,
            'name':     name,
            'months':   months_str,
            'expected_charges': [], 'expected_credits': [],
            'found_charges':    [], 'found_credits':    [],
            'missing_charges':  [], 'missing_credits':  [],
            'unexpected_pdf':   [{'kind': e['kind'], 'month': e['month_str'],
                                  'cost': e['cost']} for e in entries],
            'cost':     sum(abs(e['cost']) for e in entries),
            'plan':     plan,
            'reason':   f"Member not found in Excel roster — PDF has {type_label.lower()} entries",
        })

    # Sort: incorrect first, then missing, then ok, then no_adj_needed
    sort_order = {'incorrect': 0, 'missing': 1, 'ok': 2, 'no_adj_needed': 3}
    validations.sort(key=lambda v: (sort_order[v['status']], v['name']))

    return validations
