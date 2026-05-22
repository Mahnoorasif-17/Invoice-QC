"""
QC Engine — Excel is the source of truth.

Two reports:

  1. BILLING ISSUES (Rules 1-4) — walk Excel rows
       Rule 1: Terminated but billed
       Rule 2: Active but not billed (start month ≤ invoice month)
       Rule 3: Future start (in a month AFTER invoice month) but billed
       Rule 4: SKU mismatch (ERROR) / Same SKU + different amount (PEPM warning)

  2. ADJUSTMENTS & CREDITS — one row per (member, plan) combination
       Walks BOTH Excel members AND PDF entries; whichever side has data,
       a row appears.
       Status values:
         ok               PDF matches Excel expectation
         missing          Excel predicts adjustment but PDF doesn't have it
         incorrect        PDF entry contradicts Excel (wrong member, before
                          start, after end, etc.)
         no_adj_needed    Excel member is in data window but needs no adj
         unexplained      PDF has entry, but Excel doesn't explain it
                          (no end date for credit, no enrollment in window
                          for charge, etc.) — needs human review

Plan-change detection: if a member has BOTH a charge and a credit in PDF for
the SAME coverage month but DIFFERENT plan codes, that's recognized as a
plan-change pattern and the credit (old plan) + charge (new plan) are linked.

Key rules:
  * EndDate in any past month → NO credit needed for that month
    (any day = some insurance used = full month billed)
    Credits are only needed for months billed AFTER end month.
  * Start in any past month + EnrolledOn in data window →
    charge needed for every month from start through invoice_month-1.
  * Start IN invoice month → no charge needed (regular billing covers).
  * End IN invoice month → no credit needed (regular billing covers).
"""

import calendar
import re
from datetime import datetime
from typing import Optional, Any, List, Dict, Tuple
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


def normalize_plan(s: Any) -> str:
    return re.sub(r'[^A-Z0-9]', '', str(s or '').upper())


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
# MAIN ENTRY POINT
# ────────────────────────────────────────────────────────────────────
def run_qc(rows: list, pdf_data: dict, ctx: dict) -> dict:
    # ── BILLING RULES 1-4 ───────────────────────────────────────────
    member_issues = _run_billing_rules(rows, ctx)

    # ── LATE ADJUSTMENTS (>60 days) ─────────────────────────────────
    late_adj = _collect_late_adjustments(pdf_data, ctx)

    # ── ADJUSTMENT/CREDIT VALIDATION (per-member) ───────────────────
    validations = validate_adjustments(rows, pdf_data, ctx)

    # ── ROLL-UP STATUS ──────────────────────────────────────────────
    has_errors = any(any(i['sev'] == 'error' for i in m['iss']) for m in member_issues)
    has_warnings = any(any(i['sev'] in ('warning', 'pepm', 'approval') for i in m['iss'])
                       for m in member_issues)
    if any(v['status'] == 'incorrect' for v in validations):
        has_errors = True
    if any(v['status'] in ('missing', 'unexplained') for v in validations):
        has_warnings = True

    return {
        'groupId':              pdf_data.get('groupId', '') or '',
        'groupName':            pdf_data.get('groupName', '') or '',
        'invoiceAmount':        pdf_data.get('invoiceAmount', 0) or 0,
        'currentPeriodAmount':  pdf_data.get('currentPeriodAmount', 0) or 0,
        'adjCharges':           pdf_data.get('adjustmentChargesTotal', 0) or 0,
        'adjCredits':           pdf_data.get('adjustmentCreditsTotal', 0) or 0,
        'mi':                   member_issues,
        'la':                   late_adj,
        'validations':          validations,
        'hasErrors':            has_errors,
        'hasWarnings':          not has_errors and (has_warnings or len(late_adj) > 0),
        'isClean':              not has_errors and not has_warnings and len(late_adj) == 0,
    }


# ────────────────────────────────────────────────────────────────────
# BILLING RULES 1-4
# ────────────────────────────────────────────────────────────────────
def _run_billing_rules(rows: list, ctx: dict) -> list:
    member_issues = []
    for row in rows:
        fn = str(row.get('FirstName') or '').strip()
        ln = str(row.get('LastName') or '').strip()
        name = f"{fn} {ln}".strip() or '(no name)'

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
            member_issues.append({'name': name, 'iss': issues})
    return member_issues


def _collect_late_adjustments(pdf_data: dict, ctx: dict) -> list:
    """List PDF adjustments older than the 60-day boundary."""
    late = []
    for r in (pdf_data.get('adjustmentChargeRoster') or []):
        d = parse_month_string(r.get('coverageMonth'))
        if d and d < ctx['boundary60']:
            late.append({
                'name': f"{r.get('firstName', '')} {r.get('lastName', '')}".strip(),
                'type': 'Charge', 'month': r.get('coverageMonth', ''),
                'cost': to_float(r.get('cost')) or 0, 'plan': r.get('planCode', '')})
    for r in (pdf_data.get('adjustmentCreditRoster') or []):
        d = parse_month_string(r.get('coverageMonth'))
        if d and d < ctx['boundary60']:
            late.append({
                'name': f"{r.get('firstName', '')} {r.get('lastName', '')}".strip(),
                'type': 'Credit', 'month': r.get('coverageMonth', ''),
                'cost': to_float(r.get('cost')) or 0, 'plan': r.get('planCode', '')})
    return late


# ────────────────────────────────────────────────────────────────────
# ADJUSTMENT / CREDIT VALIDATION — comprehensive per-member
# ────────────────────────────────────────────────────────────────────
def validate_adjustments(rows: list, pdf_data: dict, ctx: dict) -> list:
    """
    Walks BOTH Excel (in scope) AND PDF entries. Produces one row per member,
    each with: status, type label, list of months, plan(s), cost, reason.
    """
    # ── Index Excel by normalized name ──────────────────────────────
    excel_by_name: Dict[str, dict] = {}
    for row in rows:
        fn = str(row.get('FirstName') or '').strip()
        ln = str(row.get('LastName') or '').strip()
        nk = normalize_name(fn + ln)
        if nk:
            excel_by_name[nk] = row

    # ── Index PDF entries by normalized name ────────────────────────
    pdf_charges: Dict[str, list] = {}
    pdf_credits: Dict[str, list] = {}

    for r in (pdf_data.get('adjustmentChargeRoster') or []):
        fn = str(r.get('firstName', '')).strip()
        ln = str(r.get('lastName', '')).strip()
        nk = normalize_name(fn + ln)
        pdf_charges.setdefault(nk, []).append({
            'month_dt':  parse_month_string(r.get('coverageMonth')),
            'month_str': str(r.get('coverageMonth', '') or '?'),
            'plan':      str(r.get('planCode', '') or ''),
            'cost':      to_float(r.get('cost')) or 0,
            'first':     fn, 'last': ln, 'raw': r,
        })

    for r in (pdf_data.get('adjustmentCreditRoster') or []):
        fn = str(r.get('firstName', '')).strip()
        ln = str(r.get('lastName', '')).strip()
        nk = normalize_name(fn + ln)
        pdf_credits.setdefault(nk, []).append({
            'month_dt':  parse_month_string(r.get('coverageMonth')),
            'month_str': str(r.get('coverageMonth', '') or '?'),
            'plan':      str(r.get('planCode', '') or ''),
            'cost':      to_float(r.get('cost')) or 0,
            'first':     fn, 'last': ln, 'raw': r,
        })

    # Build the union of all member keys (Excel-in-scope ∪ PDF members)
    all_keys = set()

    # Add Excel members whose EnrolledOn or EndedOn falls in data window
    for nk, row in excel_by_name.items():
        eo = parse_date(row.get('EnrolledOn'))
        on = parse_date(row.get('EndedOn'))
        if (eo and in_window(eo, ctx['windowStart'], ctx['windowEnd'])) or \
           (on and in_window(on, ctx['windowStart'], ctx['windowEnd'])):
            all_keys.add(nk)

    # Add anyone with PDF entries
    all_keys.update(pdf_charges.keys())
    all_keys.update(pdf_credits.keys())

    # Validate each member
    validations = []
    for nk in all_keys:
        validations.append(_validate_one_member(
            nk, excel_by_name.get(nk),
            pdf_charges.get(nk, []), pdf_credits.get(nk, []),
            ctx,
        ))

    # Sort: incorrect → missing → unexplained → ok → no_adj_needed, then by name
    sort_order = {'incorrect': 0, 'missing': 1, 'unexplained': 2,
                  'ok': 3, 'no_adj_needed': 4}
    validations.sort(key=lambda v: (sort_order.get(v['status'], 9), v['name']))
    return validations


def _validate_one_member(
    nk: str,
    excel_row: Optional[dict],
    pdf_chgs: list,
    pdf_crds: list,
    ctx: dict,
) -> dict:
    """Build one validation row for a single member (by name)."""

    # ── 1. Extract Excel facts ─────────────────────────────────────
    if excel_row is not None:
        fn = str(excel_row.get('FirstName') or '').strip()
        ln = str(excel_row.get('LastName') or '').strip()
        name = f"{fn} {ln}".strip() or '(no name)'
        st = str(excel_row.get('EmploymentStatus') or '').strip()
        sd = parse_date(excel_row.get('StartDate'))
        eo = parse_date(excel_row.get('EnrolledOn'))
        ed = parse_date(excel_row.get('EndDate'))
        on = parse_date(excel_row.get('EndedOn'))
        excel_plan = str(excel_row.get('CarrierPlanCode') or '').strip()
        plan_cost = to_float(excel_row.get('PlanCost')) or 0
    else:
        # PDF-only member — use the name from PDF entries
        sample = (pdf_chgs + pdf_crds)[0]
        fn = sample['first']
        ln = sample['last']
        name = f"{fn} {ln}".strip() or '(no name)'
        st, sd, eo, ed, on = '', None, None, None, None
        excel_plan = ''
        plan_cost = 0

    eo_in_window = eo is not None and in_window(eo, ctx['windowStart'], ctx['windowEnd'])
    on_in_window = on is not None and in_window(on, ctx['windowStart'], ctx['windowEnd'])

    # ── 2. Build expected charge / credit months from Excel ────────
    expected_charge_months: List[datetime] = []
    if eo_in_window and st == 'Active' and sd:
        sd_month = month_floor(sd)
        if sd_month < ctx['invoiceStart']:
            for m in months_between(sd_month, ctx['invoiceStart']):
                expected_charge_months.append(m)

    expected_credit_months: List[datetime] = []
    if on_in_window and ed:
        ed_month = month_floor(ed)
        first_credit = next_month(ed_month)
        if first_credit < ctx['invoiceStart']:
            for m in months_between(first_credit, ctx['invoiceStart']):
                expected_credit_months.append(m)

    # ── 3. Match PDF entries against expectations ──────────────────
    # We match by (month_dt) — plan code may differ if there was a plan change.
    matched_charges_idx = set()
    matched_credits_idx = set()

    found_charges: List[datetime] = []
    missing_charges: List[datetime] = []
    for em in expected_charge_months:
        idx = None
        for i, pc in enumerate(pdf_chgs):
            if i in matched_charges_idx:
                continue
            if pc['month_dt'] == em:
                idx = i
                matched_charges_idx.add(i)
                break
        (found_charges if idx is not None else missing_charges).append(em)

    found_credits: List[datetime] = []
    missing_credits: List[datetime] = []
    for em in expected_credit_months:
        idx = None
        for i, pc in enumerate(pdf_crds):
            if i in matched_credits_idx:
                continue
            if pc['month_dt'] == em:
                idx = i
                matched_credits_idx.add(i)
                break
        (found_credits if idx is not None else missing_credits).append(em)

    # Unmatched PDF entries
    unmatched_chg = [pc for i, pc in enumerate(pdf_chgs) if i not in matched_charges_idx]
    unmatched_crd = [pc for i, pc in enumerate(pdf_crds) if i not in matched_credits_idx]

    # ── 4. Plan-change detection ───────────────────────────────────
    # Pair an unmatched charge with an unmatched credit if they share the
    # SAME coverage month and have DIFFERENT plan codes.
    plan_changes: List[Tuple[dict, dict]] = []
    used_chg_idx, used_crd_idx = set(), set()
    for i, chg in enumerate(unmatched_chg):
        for j, crd in enumerate(unmatched_crd):
            if j in used_crd_idx:
                continue
            if chg['month_dt'] and crd['month_dt'] and chg['month_dt'] == crd['month_dt'] \
                    and normalize_plan(chg['plan']) != normalize_plan(crd['plan']):
                plan_changes.append((chg, crd))
                used_chg_idx.add(i)
                used_crd_idx.add(j)
                break

    leftover_chg = [c for i, c in enumerate(unmatched_chg) if i not in used_chg_idx]
    leftover_crd = [c for i, c in enumerate(unmatched_crd) if i not in used_crd_idx]

    # ── 5. Classify each PDF entry vs Excel timeline ───────────────
    incorrect_reasons: List[str] = []
    unexplained_reasons: List[str] = []

    for chg in leftover_chg:
        m = chg['month_dt']
        if m is None:
            unexplained_reasons.append(f"Charge with unparseable month {chg['month_str']!r}")
            continue
        # If member started after this month → INCORRECT
        if sd and m < month_floor(sd):
            incorrect_reasons.append(
                f"Charge for {format_month(m)} but member started {sd.strftime('%m/%d/%Y')}")
            continue
        # If member ended before this month → INCORRECT
        if ed and m > month_floor(ed):
            incorrect_reasons.append(
                f"Charge for {format_month(m)} but member ended {ed.strftime('%m/%d/%Y')}")
            continue
        # Within employment but Excel didn't predict (no EnrolledOn in window)
        unexplained_reasons.append(
            f"Charge for {format_month(m)} on {chg['plan']} — not predicted by Excel data window")

    for crd in leftover_crd:
        m = crd['month_dt']
        if m is None:
            unexplained_reasons.append(f"Credit with unparseable month {crd['month_str']!r}")
            continue
        if sd and m < month_floor(sd):
            incorrect_reasons.append(
                f"Credit for {format_month(m)} but member started {sd.strftime('%m/%d/%Y')}")
            continue
        # Credit but no end date in Excel
        if not ed:
            unexplained_reasons.append(
                f"Credit for {format_month(m)} on {crd['plan']} but member has no EndDate in Excel")
            continue
        # Credit for end-month itself → INCORRECT (we don't owe credit for the end month)
        if m == month_floor(ed):
            incorrect_reasons.append(
                f"Credit for {format_month(m)} on {crd['plan']} — member ended "
                f"{ed.strftime('%m/%d/%Y')}, that month was used (no credit owed)")
            continue
        # Credit for a month before end → INCORRECT
        if m < month_floor(ed):
            incorrect_reasons.append(
                f"Credit for {format_month(m)} but member ended {ed.strftime('%m/%d/%Y')}")
            continue
        unexplained_reasons.append(
            f"Credit for {format_month(m)} on {crd['plan']} — not predicted by Excel data window")

    # Plan changes: validated against employment dates
    plan_change_notes: List[str] = []
    for chg, crd in plan_changes:
        m = chg['month_dt']
        bad = False
        if sd and m and m < month_floor(sd):
            incorrect_reasons.append(
                f"Plan change for {format_month(m)} but member started {sd.strftime('%m/%d/%Y')}")
            bad = True
        if ed and m and m > month_floor(ed):
            incorrect_reasons.append(
                f"Plan change for {format_month(m)} but member ended {ed.strftime('%m/%d/%Y')}")
            bad = True
        if not bad:
            plan_change_notes.append(
                f"Plan change for {format_month(m)}: {crd['plan']} → {chg['plan']} "
                f"(credit ${abs(crd['cost']):,.2f}, charge ${chg['cost']:,.2f})")

    # ── 6. Build verdict + labels ──────────────────────────────────
    type_parts: List[str] = []
    if expected_charge_months or pdf_chgs:
        type_parts.append('Charge')
    if expected_credit_months or pdf_crds:
        type_parts.append('Credit')
    if plan_changes:
        if 'Plan Change' not in type_parts:
            type_parts.append('Plan Change')
    type_label = ' & '.join(type_parts) if type_parts else 'None'

    # Build month list (deduplicated, in PDF order then expected order)
    # Normalize all months to "Mon YYYY" format for consistent dedup
    months_seen = []
    def _add_month(s, dt=None):
        # Prefer parsed datetime form for consistency
        if dt is not None:
            label = format_month(dt)
        else:
            parsed = parse_month_string(s)
            label = format_month(parsed) if parsed else s
        if label and label not in months_seen:
            months_seen.append(label)
    for pc in pdf_chgs: _add_month(pc['month_str'], pc['month_dt'])
    for pc in pdf_crds: _add_month(pc['month_str'], pc['month_dt'])
    for em in expected_charge_months: _add_month(None, em)
    for em in expected_credit_months: _add_month(None, em)
    months_str = ', '.join(months_seen) if months_seen else '—'

    # Plans list (deduplicated)
    plans_seen = []
    for pc in (pdf_chgs + pdf_crds):
        p = pc['plan']
        if p and p not in plans_seen:
            plans_seen.append(p)
    if excel_plan and excel_plan not in plans_seen:
        plans_seen.append(excel_plan)
    plan_str = ' / '.join(plans_seen) if plans_seen else (excel_plan or '')

    # Status decision tree
    has_expected = bool(expected_charge_months or expected_credit_months)
    has_missing  = bool(missing_charges or missing_credits)
    has_incorrect = bool(incorrect_reasons)
    has_unexplained = bool(unexplained_reasons)

    if has_incorrect:
        status = 'incorrect'
        parts = incorrect_reasons[:]
        if has_missing:
            mp = []
            if missing_charges:
                mp.append(f"Missing Charge(s) for {', '.join(format_month(m) for m in missing_charges)}")
            if missing_credits:
                mp.append(f"Missing Credit(s) for {', '.join(format_month(m) for m in missing_credits)}")
            parts.extend(mp)
        if plan_change_notes:
            parts.extend(plan_change_notes)
        reason = '; '.join(parts)

    elif has_missing:
        status = 'missing'
        parts = []
        if missing_charges:
            parts.append(f"Missing Charge(s) for {', '.join(format_month(m) for m in missing_charges)}")
        if missing_credits:
            parts.append(f"Missing Credit(s) for {', '.join(format_month(m) for m in missing_credits)}")
        approval_months = [m for m in (missing_charges + missing_credits) if m < ctx['boundary60']]
        if approval_months:
            parts.append("NEEDS APPROVAL — older than 60 days")
        if plan_change_notes:
            parts.extend(plan_change_notes)
        if has_unexplained:
            parts.extend(unexplained_reasons)
        reason = '; '.join(parts)

    elif has_unexplained:
        status = 'unexplained'
        parts = unexplained_reasons[:]
        if plan_change_notes:
            parts.extend(plan_change_notes)
        reason = '; '.join(parts)

    elif has_expected:
        status = 'ok'
        parts = []
        if eo_in_window: parts.append(f"enrolled {eo.strftime('%m/%d/%Y')}")
        if on_in_window: parts.append(f"entered EN {on.strftime('%m/%d/%Y')}")
        if sd: parts.append(f"start {sd.strftime('%m/%d/%Y')}")
        if ed: parts.append(f"end {ed.strftime('%m/%d/%Y')}")
        if plan_change_notes:
            reason = "All expected adjustments found in PDF (" + ', '.join(parts) + "); " + \
                     '; '.join(plan_change_notes)
        else:
            reason = "All expected adjustments found in PDF (" + ', '.join(parts) + ")"

    elif plan_change_notes:
        # Pure plan-change scenario — credit + charge same month, different plans
        status = 'ok'
        reason = '; '.join(plan_change_notes)

    else:
        # Excel-in-scope but no expectation, AND no PDF entries leftover → NO ADJ NEEDED
        status = 'no_adj_needed'
        reasons = []
        if eo_in_window and (not sd or month_floor(sd) >= ctx['invoiceStart']):
            if sd:
                reasons.append(f"enrolled {eo.strftime('%m/%d/%Y')} but start "
                               f"{sd.strftime('%m/%d/%Y')} is in invoice month — "
                               f"regular billing covers")
            else:
                reasons.append(f"enrolled {eo.strftime('%m/%d/%Y')} but no start date")
        if on_in_window and ed and month_floor(ed) < ctx['invoiceStart']:
            reasons.append(f"ended {ed.strftime('%m/%d/%Y')} — used insurance that month, "
                           f"no credit needed")
        if on_in_window and ed and month_floor(ed) >= ctx['invoiceStart']:
            reasons.append(f"ended {ed.strftime('%m/%d/%Y')} — in invoice month, "
                           f"regular billing covers")
        reason = '; '.join(reasons) if reasons else 'In data window but no adjustments needed'

    # ── 7. Compute total amount (charges positive, credits negative) ─
    total_amount = 0
    for pc in pdf_chgs: total_amount += pc['cost']
    for pc in pdf_crds: total_amount += pc['cost']  # already signed
    # If no PDF entries but expected, use plan_cost × predicted months
    if not pdf_chgs and not pdf_crds and (expected_charge_months or expected_credit_months):
        if expected_charge_months:
            total_amount = plan_cost * len(expected_charge_months)
        if expected_credit_months:
            total_amount = -plan_cost * len(expected_credit_months)

    return {
        'status':           status,
        'type':             type_label,
        'name':             name,
        'months':           months_str,
        'plan':             plan_str,
        'cost':             total_amount,
        'reason':           reason,
        # detail for export
        'expected_charges': [format_month(m) for m in expected_charge_months],
        'expected_credits': [format_month(m) for m in expected_credit_months],
        'found_charges':    [format_month(m) for m in found_charges],
        'found_credits':    [format_month(m) for m in found_credits],
        'missing_charges':  [format_month(m) for m in missing_charges],
        'missing_credits':  [format_month(m) for m in missing_credits],
        'plan_changes':     plan_change_notes,
        'pdf_entries': [
            {'type': 'Charge', 'month': pc['month_str'], 'plan': pc['plan'], 'cost': pc['cost']}
            for pc in pdf_chgs
        ] + [
            {'type': 'Credit', 'month': pc['month_str'], 'plan': pc['plan'], 'cost': pc['cost']}
            for pc in pdf_crds
        ],
    }
