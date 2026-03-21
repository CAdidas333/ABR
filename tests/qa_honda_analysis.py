"""
QA Deep Analysis — Honda Feb 2026 Matching Engine Results

Performs detailed spot-checks and false-positive detection:
  a) Spot-check Rule 0 (check# confirmed) matches
  b) Spot-check Rule 1 (unique amount) matches — look for order-dependency
  c) Analyze 4 false-positive outstanding deposits
  d) Check for reversed-sign matches
  e) Verify 3 near-amount matches (55% confidence)
  f) Check for duplicate assignments (bank or DMS used twice)
  g) Analyze 15 medium-confidence matches (60-84%)

IMPORTANT: The Feb DMS and Jan DMS parse with overlapping transaction IDs
(both start at 1). The engine runs two separate run_matching() calls, so
matches from the first call reference Feb DMS IDs and matches from the
prior-period call reference Jan DMS IDs. We must use separate lookups.
"""

import sys
import os
import random
from collections import Counter, defaultdict
from datetime import date

sys.path.insert(0, os.path.dirname(__file__))

from matching_engine import (
    parse_bofa_bai_csv,
    parse_dms_xlsx,
    run_full_reconciliation,
    run_matching,
    MatchConfig,
    Transaction,
    _get_dms_check,
)

# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------
DOCS = os.path.join(os.path.dirname(__file__), '..', 'docs')
BANK_FILE = os.path.join(DOCS, 'Honda Bkstmt.csv')
FEB_GL_FILE = os.path.join(DOCS, 'honda Feb GL.xlsx')
JAN_GL_FILE = os.path.join(DOCS, 'honda Jan GL.xlsx')

# Ground truth: outstanding deposits from bank rec
OUTSTANDING_DEPOSITS = {
    '022626MVDA': 32098.32,
    '022726FI': 82583.49,
    '022726MVDA': 30130.01,
    '022726PSC': 258.98,
    '022826FI': 75744.00,
    '022826MVDA': 53398.76,
}
CORRECTLY_UNMATCHED = {'022626MVDA', '022726FI'}
FALSE_POSITIVE_REFS = {k for k in OUTSTANDING_DEPOSITS if k not in CORRECTLY_UNMATCHED}


def section_header(title):
    print(f"\n{'=' * 78}")
    print(f"  {title}")
    print(f"{'=' * 78}")


def subsection(title):
    print(f"\n  --- {title} ---")


def main():
    random.seed(42)  # Reproducible spot-checks

    print("=" * 78)
    print("  QA DEEP ANALYSIS — Honda Feb 2026 Matching Engine")
    print("=" * 78)

    # --- Load data ---
    print("\nLoading data...")
    bank_txns = parse_bofa_bai_csv(BANK_FILE)
    feb_dms = parse_dms_xlsx(FEB_GL_FILE)
    jan_dms = parse_dms_xlsx(JAN_GL_FILE)
    print(f"  Bank: {len(bank_txns)}, Feb DMS: {len(feb_dms)}, Jan DMS: {len(jan_dms)}")

    # ID collision warning: both Feb and Jan DMS start IDs at 1
    feb_ids = {t.transaction_id for t in feb_dms}
    jan_ids = {t.transaction_id for t in jan_dms}
    overlap = feb_ids & jan_ids
    print(f"  WARNING: Feb/Jan DMS ID overlap: {len(overlap)} IDs "
          f"(IDs 1-{min(len(feb_dms), len(jan_dms))})")
    print(f"  Using separate lookups for Feb vs Jan DMS to avoid collisions.")

    # Build separate lookup dicts
    bank_by_id = {t.transaction_id: t for t in bank_txns}
    feb_dms_by_id = {t.transaction_id: t for t in feb_dms}
    jan_dms_by_id = {t.transaction_id: t for t in jan_dms}

    # --- Run engine ---
    config = MatchConfig()
    results = run_full_reconciliation(bank_txns, feb_dms, config, prior_dms_txns=jan_dms)
    all_matches = results['all_matches']
    one_to_one = results['one_to_one_matches']
    prior_matches = results['prior_period_matches']
    cvr_matches = results['cvr_matches']
    print(f"  Total matches: {len(all_matches)}")
    print(f"    1:1 current: {len(one_to_one)}, prior: {len(prior_matches)}, CVR: {len(cvr_matches)}")

    def get_dms_txn(match):
        """Look up DMS transaction using correct dict based on match type."""
        did = match.dms_transaction_ids[0]
        if match.match_type == "PRIOR_PERIOD":
            return jan_dms_by_id.get(did)
        else:
            return feb_dms_by_id.get(did)

    def get_dms_txns_for_match(match):
        """Look up all DMS transactions for a match."""
        dms_dict = jan_dms_by_id if match.match_type == "PRIOR_PERIOD" else feb_dms_by_id
        return [dms_dict.get(did) for did in match.dms_transaction_ids]

    # =====================================================================
    # (a) SPOT-CHECK RULE 0 MATCHES (CHECK# CONFIRMED)
    # =====================================================================
    section_header("(a) SPOT-CHECK: Rule 0 — Check# Confirmed Matches")

    rule0_matches = [m for m in all_matches
                     if 'Check#' in m.score_breakdown and 'confirmed' in m.score_breakdown]
    print(f"\n  Total Rule 0 matches: {len(rule0_matches)}")

    sample_size = min(25, len(rule0_matches))
    sample_r0 = random.sample(rule0_matches, sample_size)

    failures_r0 = []
    print(f"  Spot-checking {sample_size} random Rule 0 matches...\n")
    for m in sample_r0:
        bid = m.bank_transaction_ids[0]
        bank = bank_by_id.get(bid)
        dms = get_dms_txn(m)

        if not bank or not dms:
            failures_r0.append((m, "MISSING TXN",
                f"Bank ID={bid} found={bank is not None}, "
                f"DMS ID={m.dms_transaction_ids[0]} found={dms is not None} "
                f"(type={m.match_type})"))
            continue

        # Verify check numbers actually match
        bank_ck = bank.check_number
        dms_ck = _get_dms_check(dms)
        ck_match = (bank_ck == dms_ck) and bool(bank_ck)

        # Verify amounts are EXACTLY identical
        amt_match = bank.amount == dms.amount

        # Cross-check with stored amounts on the match
        stored_match = (m.bank_amount == m.dms_amount and m.amount_difference == 0.0)

        status = "OK" if (ck_match and amt_match) else "FAIL"
        if status == "FAIL":
            reason = []
            if not ck_match:
                reason.append(f"check# mismatch: bank='{bank_ck}' dms='{dms_ck}'")
            if not amt_match:
                reason.append(f"amount mismatch: bank={bank.amount} dms={dms.amount}")
            failures_r0.append((m, status, '; '.join(reason)))

        period_tag = "JAN" if m.match_type == "PRIOR_PERIOD" else "FEB"
        print(f"    Match #{m.match_id:>4} [{period_tag}]: Check #{bank_ck:<8} "
              f"Bank=${bank.amount:>12,.2f} DMS=${dms.amount:>12,.2f} "
              f"CK={'MATCH' if ck_match else 'MISMATCH':>8} "
              f"AMT={'EXACT' if amt_match else 'DIFF':>5} -> {status}")

    if failures_r0:
        print(f"\n  ** {len(failures_r0)} FAILURES in Rule 0 spot-check: **")
        for m, status, reason in failures_r0:
            print(f"     Match #{m.match_id}: {reason}")
    else:
        print(f"\n  ALL {sample_size} Rule 0 spot-checks PASSED.")
        print(f"  Check numbers genuinely match, amounts are exact (zero difference).")

    # =====================================================================
    # (b) SPOT-CHECK RULE 1 MATCHES (UNIQUE AMOUNT)
    # =====================================================================
    section_header("(b) SPOT-CHECK: Rule 1 — Unique Amount Matches (Order-Dependency Check)")

    rule1_matches = [m for m in all_matches
                     if 'Unique amount' in m.score_breakdown]
    print(f"\n  Total Rule 1 matches: {len(rule1_matches)}")

    # For order-dependency check, count candidates in the CORRECT DMS pool
    # Current-period Rule 1 matches use Feb DMS; prior-period use Jan DMS
    sample_size_r1 = min(25, len(rule1_matches))
    sample_r1 = random.sample(rule1_matches, sample_size_r1)

    order_dependency_issues = []
    print(f"  Spot-checking {sample_size_r1} random Rule 1 matches for order-dependency...\n")

    for m in sample_r1:
        bid = m.bank_transaction_ids[0]
        bank = bank_by_id.get(bid)
        dms = get_dms_txn(m)
        if not bank or not dms:
            continue

        # Count ALL DMS candidates at exact amount in the appropriate pool
        # Rule 1 runs AFTER Rule 0 has consumed some. We check the FULL pool.
        if m.match_type == "PRIOR_PERIOD":
            pool = jan_dms
            pool_label = "JAN"
        else:
            pool = feb_dms
            pool_label = "FEB"

        all_candidates_full = [d for d in pool if d.amount == bank.amount]
        n_full = len(all_candidates_full)

        # Was this actually unique in the full list?
        if n_full > 1:
            consumed = [d for d in all_candidates_full
                        if d.transaction_id != dms.transaction_id and d.is_matched]
            unconsumed = [d for d in all_candidates_full
                          if d.transaction_id != dms.transaction_id and not d.is_matched]
            order_dependency_issues.append({
                'match': m,
                'bank': bank,
                'dms': dms,
                'full_count': n_full,
                'consumed_count': len(consumed),
                'unconsumed_remaining': len(unconsumed),
                'pool': pool_label,
            })

        unique_str = f"{'YES' if n_full == 1 else 'NO':>3} (full={n_full})"
        print(f"    Match #{m.match_id:>4} [{pool_label}]: Amount=${bank.amount:>12,.2f} "
              f"DMS candidates (full {pool_label} pool)={n_full:>2} "
              f"Truly unique={unique_str} "
              f"Conf={m.confidence_score:.0f}%")

    if order_dependency_issues:
        print(f"\n  ** {len(order_dependency_issues)} ORDER-DEPENDENCY cases found: **")
        print(f"  These became 'unique' only because earlier phases consumed other candidates.\n")
        for issue in order_dependency_issues:
            m = issue['match']
            b = issue['bank']
            print(f"    Match #{m.match_id} [{issue['pool']}]: ${b.amount:>12,.2f} "
                  f"had {issue['full_count']} candidates originally, "
                  f"{issue['consumed_count']} consumed by earlier matches, "
                  f"{issue['unconsumed_remaining']} still unmatched")
            # Note: this is expected — Rule 0 consuming check matches before
            # Rule 1 runs is by design. It only becomes problematic if the
            # WRONG candidate was consumed.
    else:
        print(f"\n  ALL {sample_size_r1} Rule 1 matches were truly unique in their DMS pool.")

    # =====================================================================
    # (c) ANALYZE FALSE-POSITIVE OUTSTANDING DEPOSITS
    # =====================================================================
    section_header("(c) FALSE-POSITIVE OUTSTANDING DEPOSITS")

    print(f"\n  Outstanding deposits from bank rec: {len(OUTSTANDING_DEPOSITS)}")
    print(f"  Correctly remained unmatched: {sorted(CORRECTLY_UNMATCHED)}")
    print(f"  Allegedly false-positive (matched but shouldn't be): {sorted(FALSE_POSITIVE_REFS)}")

    subsection("Checking each outstanding deposit against engine results")

    fp_confirmed = 0
    fp_not_confirmed = 0

    for ref, expected_amount in sorted(OUTSTANDING_DEPOSITS.items()):
        is_fp_expected = ref in FALSE_POSITIVE_REFS
        label = "EXPECTED FP" if is_fp_expected else "EXPECTED UNMATCHED"

        print(f"\n  Outstanding Deposit: {ref} = ${expected_amount:,.2f} [{label}]")

        # Find matching DMS transaction(s) by reference AND amount in Feb
        dms_by_ref = [d for d in feb_dms
                      if d.reference_number == ref]
        dms_by_amt = [d for d in feb_dms
                      if abs(d.amount - expected_amount) < 0.005]

        # Combine candidates
        all_candidates = {}
        for d in dms_by_ref + dms_by_amt:
            all_candidates[d.transaction_id] = d

        if not all_candidates:
            print(f"    -> No DMS transaction found in Feb GL")
            continue

        for did, dms in sorted(all_candidates.items()):
            print(f"    DMS ID={did}: ${dms.amount:,.2f} date={dms.transaction_date} "
                  f"ref='{dms.reference_number}' type={dms.type_code} "
                  f"matched={dms.is_matched}")

            if dms.is_matched:
                # Find which match consumed this DMS txn
                # Must search in current-period matches (not prior)
                consuming_match = None
                for m in one_to_one + cvr_matches:
                    if did in m.dms_transaction_ids:
                        consuming_match = m
                        break

                if consuming_match:
                    fp_confirmed += 1
                    for mbid in consuming_match.bank_transaction_ids:
                        mb = bank_by_id.get(mbid)
                        if mb:
                            days_gap = abs((mb.transaction_date - dms.transaction_date).days)
                            print(f"      MATCHED -> Match #{consuming_match.match_id} "
                                  f"(conf={consuming_match.confidence_score:.0f}%)")
                            print(f"      Rule: {consuming_match.score_breakdown}")
                            print(f"      Bank txn ID={mbid}: ${mb.amount:>12,.2f} "
                                  f"date={mb.transaction_date} type={mb.type_code} "
                                  f"desc='{mb.description[:60]}'")
                            print(f"      Date gap: {days_gap} days")

                            # How many bank/DMS at this amount?
                            same_amt_bank = [b for b in bank_txns
                                             if abs(b.amount - expected_amount) < 0.005]
                            same_amt_dms = [d2 for d2 in feb_dms
                                            if abs(d2.amount - expected_amount) < 0.005]
                            print(f"      Bank txns at ${expected_amount:,.2f}: {len(same_amt_bank)}")
                            print(f"      Feb DMS txns at ${expected_amount:,.2f}: {len(same_amt_dms)}")

                    if is_fp_expected:
                        print(f"      ** CONFIRMED FALSE POSITIVE: This deposit is "
                              f"outstanding per bank rec but engine matched it.")
                else:
                    # matched but not found in current-period matches -- check prior
                    for m in prior_matches:
                        if did in m.dms_transaction_ids:
                            consuming_match = m
                            break
                    if consuming_match:
                        print(f"      NOTE: Matched in PRIOR PERIOD pass (unusual for a deposit)")
            else:
                fp_not_confirmed += 1
                print(f"      -> NOT matched (correctly outstanding)")
                if is_fp_expected:
                    print(f"      ** NOT A FALSE POSITIVE after all: The engine correctly "
                          f"left this unmatched.")

    # Show bank-side availability for each false-positive amount
    subsection("Bank transactions at false-positive deposit amounts")
    for ref in sorted(FALSE_POSITIVE_REFS):
        amt = OUTSTANDING_DEPOSITS[ref]
        bank_at_amt = [b for b in bank_txns if abs(b.amount - amt) < 0.005]
        dms_at_amt = [d for d in feb_dms if abs(d.amount - amt) < 0.005]
        print(f"\n  {ref} (${amt:,.2f}):")
        print(f"    Bank txns at this amount: {len(bank_at_amt)}")
        for b in bank_at_amt:
            print(f"      Bank ID={b.transaction_id}: ${b.amount:,.2f} "
                  f"date={b.transaction_date} type={b.type_code} "
                  f"matched={b.is_matched} desc='{b.description[:50]}'")
        print(f"    Feb DMS txns at this amount: {len(dms_at_amt)}")
        for d in dms_at_amt:
            print(f"      DMS ID={d.transaction_id}: ${d.amount:,.2f} "
                  f"date={d.transaction_date} ref='{d.reference_number}' "
                  f"matched={d.is_matched}")

    # =====================================================================
    # (d) CHECK FOR REVERSED-SIGN MATCHES
    # =====================================================================
    section_header("(d) REVERSED-SIGN MATCH CHECK")

    reversed_sign = []
    for m in all_matches:
        if len(m.bank_transaction_ids) == 1 and len(m.dms_transaction_ids) == 1:
            bank = bank_by_id.get(m.bank_transaction_ids[0])
            dms = get_dms_txn(m)
            if bank and dms:
                if (bank.amount > 0) != (dms.amount > 0) and bank.amount != 0 and dms.amount != 0:
                    reversed_sign.append((m, bank, dms))

    # Also check using stored amounts on the MatchResult itself
    reversed_sign_stored = []
    for m in all_matches:
        if m.bank_amount != 0 and m.dms_amount != 0:
            if (m.bank_amount > 0) != (m.dms_amount > 0):
                reversed_sign_stored.append(m)

    if reversed_sign:
        print(f"\n  ** {len(reversed_sign)} REVERSED-SIGN MATCHES FOUND (by txn lookup): **")
        for m, bank, dms in reversed_sign[:10]:
            period = "JAN" if m.match_type == "PRIOR_PERIOD" else "FEB"
            print(f"    Match #{m.match_id} [{period}]: "
                  f"Bank=${bank.amount:>12,.2f} vs DMS=${dms.amount:>12,.2f} "
                  f"conf={m.confidence_score:.0f}%")
            print(f"      Stored: bank_amount=${m.bank_amount:,.2f} dms_amount=${m.dms_amount:,.2f}")
            print(f"      Rule: {m.score_breakdown[:80]}")
        if len(reversed_sign) > 10:
            print(f"    ... and {len(reversed_sign) - 10} more")
    else:
        print(f"\n  No reversed-sign matches found (by txn lookup).")

    if reversed_sign_stored:
        print(f"\n  ** {len(reversed_sign_stored)} REVERSED-SIGN based on stored match amounts: **")
        for m in reversed_sign_stored[:5]:
            print(f"    Match #{m.match_id}: stored bank=${m.bank_amount:,.2f} "
                  f"dms=${m.dms_amount:,.2f} conf={m.confidence_score:.0f}%")
    else:
        print(f"\n  No reversed-sign matches based on stored amounts either. All clean.")

    # =====================================================================
    # (e) VERIFY NEAR-AMOUNT MATCHES (55% CONFIDENCE)
    # =====================================================================
    section_header("(e) NEAR-AMOUNT MATCHES (55% confidence, $0.01 tolerance)")

    near_matches = [m for m in all_matches if m.confidence_score == 55.0]
    # Also find by score breakdown tag
    near_by_tag = [m for m in all_matches if 'NEAR AMOUNT' in m.score_breakdown]
    print(f"\n  55% confidence matches: {len(near_matches)}")
    print(f"  'NEAR AMOUNT' tag matches: {len(near_by_tag)}")

    near_to_check = near_by_tag if near_by_tag else near_matches

    for m in near_to_check:
        bid = m.bank_transaction_ids[0]
        bank = bank_by_id.get(bid)
        dms = get_dms_txn(m)

        print(f"\n    Match #{m.match_id}:")
        print(f"      Rule: {m.score_breakdown}")

        # Use stored amounts (reliable regardless of ID collision)
        diff = m.bank_amount - m.dms_amount
        abs_diff = abs(diff)
        stored_diff = abs(m.amount_difference)
        within_tolerance = abs_diff <= 0.01

        print(f"      Stored bank_amount: ${m.bank_amount:>12,.2f}")
        print(f"      Stored dms_amount:  ${m.dms_amount:>12,.2f}")
        print(f"      Stored difference:  ${m.amount_difference:+.2f}")
        print(f"      Computed difference: ${diff:+.4f} (abs=${abs_diff:.4f})")
        print(f"      Tolerance check: {'WITHIN' if within_tolerance else 'EXCEEDS'} $0.01")

        if bank and dms:
            days = abs((bank.transaction_date - dms.transaction_date).days)
            period = "JAN" if m.match_type == "PRIOR_PERIOD" else "FEB"
            print(f"      Bank: ID={bid} ${bank.amount:,.2f} date={bank.transaction_date} "
                  f"type={bank.type_code} check#='{bank.check_number}'")
            print(f"      DMS [{period}]: ID={m.dms_transaction_ids[0]} ${dms.amount:,.2f} "
                  f"date={dms.transaction_date} ref='{dms.reference_number}'")
            print(f"      Date gap: {days} days")

        if not within_tolerance:
            print(f"      ** WARNING: Stored amounts show ${abs_diff:.4f} > $0.01 tolerance! **")

    # =====================================================================
    # (f) DUPLICATE ASSIGNMENT CHECK
    # =====================================================================
    section_header("(f) DUPLICATE ASSIGNMENT CHECK")

    # Check bank side
    bank_id_usage = Counter()
    for m in all_matches:
        for bid in m.bank_transaction_ids:
            bank_id_usage[bid] += 1

    dup_bank = {bid: count for bid, count in bank_id_usage.items() if count > 1}
    if dup_bank:
        print(f"\n  ** {len(dup_bank)} BANK TRANSACTIONS assigned to multiple matches: **")
        for bid, count in sorted(dup_bank.items())[:10]:
            bank = bank_by_id.get(bid)
            matches_using = [m for m in all_matches if bid in m.bank_transaction_ids]
            print(f"    Bank ID={bid} (${bank.amount:,.2f}) used in {count} matches:")
            for m in matches_using:
                print(f"      Match #{m.match_id}: conf={m.confidence_score:.0f}% "
                      f"type={m.match_type} dms_amt=${m.dms_amount:,.2f}")
    else:
        print(f"\n  No bank transactions assigned to multiple matches. CLEAN.")

    # Check DMS side -- CRITICAL: Feb and Jan IDs overlap, so we must
    # separate by match type
    print(f"\n  DMS duplicate check (separated by period to avoid ID collisions):")

    # Current-period DMS
    current_dms_usage = Counter()
    for m in one_to_one + cvr_matches + results['split_matches']:
        for did in m.dms_transaction_ids:
            current_dms_usage[did] += 1
    dup_current_dms = {did: c for did, c in current_dms_usage.items() if c > 1}

    # Prior-period DMS
    prior_dms_usage = Counter()
    for m in prior_matches:
        for did in m.dms_transaction_ids:
            prior_dms_usage[did] += 1
    dup_prior_dms = {did: c for did, c in prior_dms_usage.items() if c > 1}

    # Cross-period: same DMS ID used in both current and prior matches
    # This is a real collision if Feb and Jan DMS both have that ID
    cross_period = set(current_dms_usage.keys()) & set(prior_dms_usage.keys())

    if dup_current_dms:
        print(f"\n  ** {len(dup_current_dms)} Feb DMS IDs used in multiple CURRENT matches: **")
        for did, count in sorted(dup_current_dms.items())[:5]:
            d = feb_dms_by_id.get(did)
            print(f"    DMS ID={did} (${d.amount:,.2f} ref='{d.reference_number}') "
                  f"used {count} times")
    else:
        print(f"    Feb DMS: No duplicates within current-period matches. CLEAN.")

    if dup_prior_dms:
        print(f"\n  ** {len(dup_prior_dms)} Jan DMS IDs used in multiple PRIOR matches: **")
        for did, count in sorted(dup_prior_dms.items())[:5]:
            d = jan_dms_by_id.get(did)
            print(f"    DMS ID={did} (${d.amount:,.2f} ref='{d.reference_number}') "
                  f"used {count} times")
    else:
        print(f"    Jan DMS: No duplicates within prior-period matches. CLEAN.")

    if cross_period:
        print(f"\n  NOTE: {len(cross_period)} DMS IDs appear in both current and prior matches.")
        print(f"  This is expected because Feb and Jan DMS use overlapping ID spaces.")
        print(f"  These are DIFFERENT transactions (Feb vs Jan) that happen to share an ID number.")
    else:
        print(f"    No cross-period ID overlap in matches.")

    total_unique_bank = len(bank_id_usage)
    print(f"\n  Summary: {total_unique_bank} unique bank txns matched, "
          f"{len(current_dms_usage)} Feb DMS + {len(prior_dms_usage)} Jan DMS "
          f"across {len(all_matches)} matches.")

    # =====================================================================
    # (g) ANALYZE MEDIUM-CONFIDENCE MATCHES (60-84%)
    # =====================================================================
    section_header("(g) MEDIUM-CONFIDENCE MATCHES (60-84%) — Most Likely to Be Wrong")

    med_matches = sorted(
        [m for m in all_matches if 60 <= m.confidence_score < 85],
        key=lambda m: m.confidence_score
    )
    print(f"\n  Total medium-confidence matches: {len(med_matches)}")

    suspicious_count = 0
    for m in med_matches:
        period = "JAN" if m.match_type == "PRIOR_PERIOD" else "FEB"
        print(f"\n  Match #{m.match_id} — {m.confidence_score:.0f}% ({m.match_type}) [{period}]")
        print(f"    Rule: {m.score_breakdown}")
        print(f"    Stored: bank=${m.bank_amount:>12,.2f} dms=${m.dms_amount:>12,.2f} "
              f"diff=${m.amount_difference:.2f}")

        # Show bank txns
        for bid in m.bank_transaction_ids:
            bank = bank_by_id.get(bid)
            if bank:
                print(f"    Bank: ID={bid} ${bank.amount:>12,.2f} date={bank.transaction_date} "
                      f"type={bank.type_code} check#='{bank.check_number}' "
                      f"desc='{bank.description[:60]}'")

        # Show DMS txns (using correct pool)
        dms_list = get_dms_txns_for_match(m)
        for i, dms in enumerate(dms_list):
            if dms:
                did = m.dms_transaction_ids[i]
                print(f"    DMS:  ID={did} [{period}] ${dms.amount:>12,.2f} "
                      f"date={dms.transaction_date} type={dms.type_code} "
                      f"ref='{dms.reference_number}' desc='{dms.description[:60]}'")

        # For 1:1 matches, show competitor analysis
        if len(m.bank_transaction_ids) == 1:
            bid = m.bank_transaction_ids[0]
            bank = bank_by_id.get(bid)
            if bank:
                dms_pool = jan_dms if m.match_type == "PRIOR_PERIOD" else feb_dms
                exact_dms = [d for d in dms_pool if d.amount == bank.amount]
                if len(exact_dms) > 1:
                    print(f"    ** {len(exact_dms)} DMS candidates at ${bank.amount:,.2f}:")
                    for d in exact_dms:
                        if d.transaction_id not in m.dms_transaction_ids:
                            print(f"       Alt: ID={d.transaction_id} "
                                  f"date={d.transaction_date} "
                                  f"ref='{d.reference_number}' "
                                  f"matched={d.is_matched}")

        # Flag suspicious patterns
        if m.date_difference > 20:
            print(f"    ** SUSPICIOUS: {m.date_difference}-day date gap is very large")
            suspicious_count += 1
        if len(m.bank_transaction_ids) == 1 and len(m.dms_transaction_ids) == 1:
            bank = bank_by_id.get(m.bank_transaction_ids[0])
            dms = get_dms_txn(m)
            if bank and dms and bank.check_number and dms.reference_number:
                dms_ck = _get_dms_check(dms)
                if dms_ck and bank.check_number != dms_ck:
                    print(f"    ** SUSPICIOUS: Check# mismatch (bank='{bank.check_number}' "
                          f"dms='{dms_ck}') — both have check#s but they differ")
                    suspicious_count += 1

    # =====================================================================
    # SUMMARY
    # =====================================================================
    section_header("QA SUMMARY")

    print(f"\n  Total matches analyzed: {len(all_matches)}")

    # (a) Summary
    print(f"\n  (a) Rule 0 (check# confirmed): {len(rule0_matches)} matches")
    if failures_r0:
        print(f"      {len(failures_r0)} FAILURES in {sample_size} spot-checks")
        print(f"      ** These indicate real check#/amount verification failures **")
    else:
        print(f"      0 failures in {sample_size} spot-checks -- ALL VERIFIED CORRECT")

    # (b) Summary
    print(f"\n  (b) Rule 1 (unique amount): {len(rule1_matches)} matches")
    print(f"      {len(order_dependency_issues)} order-dependent in {sample_size_r1} spot-checks")
    if order_dependency_issues:
        print(f"      (These became unique after Rule 0 consumed candidates -- expected behavior)")

    # (c) Summary
    print(f"\n  (c) Outstanding deposit false positives:")
    print(f"      {fp_confirmed} confirmed false positives")
    print(f"      {fp_not_confirmed} correctly left unmatched")
    all_os_dms_matched = sum(1 for d in feb_dms
                              if d.reference_number in OUTSTANDING_DEPOSITS and d.is_matched)
    print(f"      {all_os_dms_matched} of 6 outstanding deposit DMS txns were matched by engine")

    # (d) Summary
    print(f"\n  (d) Reversed-sign matches: {len(reversed_sign)} (by txn lookup)")
    print(f"      Reversed-sign by stored amounts: {len(reversed_sign_stored)}")

    # (e) Summary
    print(f"\n  (e) Near-amount matches: {len(near_to_check)}")
    near_within = sum(1 for m in near_to_check if abs(m.amount_difference) <= 0.01)
    near_exceed = sum(1 for m in near_to_check if abs(m.amount_difference) > 0.01)
    print(f"      Within $0.01 tolerance: {near_within}")
    print(f"      Exceeding tolerance: {near_exceed}")

    # (f) Summary
    print(f"\n  (f) Duplicate assignments:")
    print(f"      Bank duplicates: {len(dup_bank)}")
    print(f"      Feb DMS duplicates: {len(dup_current_dms)}")
    print(f"      Jan DMS duplicates: {len(dup_prior_dms)}")

    # (g) Summary
    print(f"\n  (g) Medium-confidence matches: {len(med_matches)}")
    print(f"      Suspicious patterns: {suspicious_count}")

    # Overall verdict
    hard_failures = len(failures_r0) + len(dup_bank) + len(dup_current_dms) + len(dup_prior_dms)
    soft_issues = len(reversed_sign_stored) + near_exceed + fp_confirmed + suspicious_count

    print(f"\n  HARD FAILURES (engine bugs): {hard_failures}")
    print(f"  SOFT ISSUES (review needed): {soft_issues}")

    if hard_failures == 0 and soft_issues == 0:
        print(f"\n  VERDICT: CLEAN. No false positives or engine bugs detected.")
    elif hard_failures == 0:
        print(f"\n  VERDICT: No engine bugs, but {soft_issues} items need manual review.")
    else:
        print(f"\n  VERDICT: {hard_failures} engine bugs found -- INVESTIGATE IMMEDIATELY.")

    print(f"\n{'=' * 78}")
    print(f"  QA Analysis Complete")
    print(f"{'=' * 78}")


if __name__ == '__main__':
    main()
