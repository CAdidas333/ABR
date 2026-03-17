"""
ABR Test Suite — 1:1 Matching Integration Tests

Tests the greedy assignment matching procedure.
"""

from datetime import date

import pytest

from matching_engine import run_matching, run_full_reconciliation, MatchConfig
from conftest import create_bank_txn, create_dms_txn, load_scenario


class TestOneToOneMatching:
    def test_single_perfect_match(self, default_config):
        """One bank + one DMS, exact match."""
        d = date(2026, 3, 1)
        bank = [create_bank_txn(1, d, -1250.00, "CHECK #4521", check_number="4521")]
        dms = [create_dms_txn(1, d, -1250.00, "CHECK PAYMENT", reference="4521",
                              type_code="CHK")]

        matches, unmatched_b, unmatched_d = run_matching(bank, dms, default_config)

        assert len(matches) == 1
        assert matches[0].confidence_score >= 95
        assert len(unmatched_b) == 0
        assert len(unmatched_d) == 0

    def test_multiple_matches_greedy_assignment(self, default_config):
        """Multiple transactions, best matches assigned first."""
        d = date(2026, 3, 1)
        bank = [
            create_bank_txn(1, d, -1000.00, "CHECK #100", check_number="100"),
            create_bank_txn(2, d, -2000.00, "CHECK #200", check_number="200"),
            create_bank_txn(3, d, 5000.00, "DEPOSIT"),
        ]
        dms = [
            create_dms_txn(1, d, -1000.00, "CHECK A", reference="100", type_code="CHK"),
            create_dms_txn(2, d, -2000.00, "CHECK B", reference="200", type_code="CHK"),
            create_dms_txn(3, d, 5000.00, "DEPOSIT"),
        ]

        matches, unmatched_b, unmatched_d = run_matching(bank, dms, default_config)

        assert len(matches) == 3
        assert len(unmatched_b) == 0
        assert len(unmatched_d) == 0

        # Verify correct pairing (check numbers should match)
        for m in matches:
            if 1 in m.bank_transaction_ids:
                assert 1 in m.dms_transaction_ids
            elif 2 in m.bank_transaction_ids:
                assert 2 in m.dms_transaction_ids

    def test_duplicate_amounts_correct_pairing(self, default_config):
        """Two transactions with same amount but different check numbers pair correctly."""
        d = date(2026, 3, 5)
        bank = [
            create_bank_txn(1, d, -2500.00, "CHECK #4530", check_number="4530"),
            create_bank_txn(2, d, -2500.00, "CHECK #4531", check_number="4531"),
        ]
        dms = [
            create_dms_txn(1, d, -2500.00, "CHECK PARTS", reference="4530", type_code="CHK"),
            create_dms_txn(2, d, -2500.00, "CHECK UTILITY", reference="4531", type_code="CHK"),
        ]

        matches, _, _ = run_matching(bank, dms, default_config)

        assert len(matches) == 2

        # Bank 1 (check 4530) should match DMS 1 (check 4530)
        bank1_match = [m for m in matches if 1 in m.bank_transaction_ids][0]
        assert 1 in bank1_match.dms_transaction_ids

        # Bank 2 (check 4531) should match DMS 2 (check 4531)
        bank2_match = [m for m in matches if 2 in m.bank_transaction_ids][0]
        assert 2 in bank2_match.dms_transaction_ids

    def test_no_match_leaves_unmatched(self, default_config):
        """Transactions with no viable match remain unmatched."""
        d = date(2026, 3, 1)
        bank = [create_bank_txn(1, d, 1000.00)]
        dms = [create_dms_txn(1, d, 2000.00)]  # Different amount

        matches, unmatched_b, unmatched_d = run_matching(bank, dms, default_config)

        assert len(matches) == 0
        assert len(unmatched_b) == 1
        assert len(unmatched_d) == 1

    def test_empty_inputs(self, default_config):
        """Empty input lists produce no matches."""
        matches, unmatched_b, unmatched_d = run_matching([], [], default_config)
        assert len(matches) == 0

    def test_one_side_empty(self, default_config):
        """One empty list, one with data → all unmatched."""
        d = date(2026, 3, 1)
        bank = [create_bank_txn(1, d, 1000.00)]

        matches, unmatched_b, unmatched_d = run_matching(bank, [], default_config)
        assert len(matches) == 0
        assert len(unmatched_b) == 1

    def test_already_matched_skipped(self, default_config):
        """Pre-matched transactions are skipped."""
        d = date(2026, 3, 1)
        bank = [create_bank_txn(1, d, 1000.00)]
        bank[0].is_matched = True
        dms = [create_dms_txn(1, d, 1000.00)]

        matches, _, unmatched_d = run_matching(bank, dms, default_config)
        assert len(matches) == 0
        assert len(unmatched_d) == 1


class TestMatchStaging:
    def test_high_confidence_staged(self, default_config):
        """High-confidence matches are staged."""
        d = date(2026, 3, 1)
        bank = [create_bank_txn(1, d, -1250.00, "CHECK #4521", check_number="4521")]
        dms = [create_dms_txn(1, d, -1250.00, "CHECK", reference="4521", type_code="CHK")]

        matches, _, _ = run_matching(bank, dms, default_config)
        assert len(matches) == 1
        assert matches[0].confidence_score >= default_config.high_confidence_threshold
        assert matches[0].status == "STAGED"

    def test_medium_confidence_staged(self, default_config):
        """Medium-confidence matches are also staged."""
        d = date(2026, 3, 1)
        bank = [create_bank_txn(1, d, 5000.00, "DEPOSIT")]
        dms = [create_dms_txn(1, date(2026, 3, 4), 5000.00, "CUSTOMER DEPOSIT")]

        matches, _, _ = run_matching(bank, dms, default_config)
        assert len(matches) == 1
        assert (default_config.medium_confidence_threshold <=
                matches[0].confidence_score <
                default_config.high_confidence_threshold)

    def test_below_threshold_not_staged(self, default_config):
        """Matches below low threshold are not staged."""
        # With default 7-day window, date 8+ days apart gives date score 0
        # Amount match only: 100 * 0.40 + 50 * 0.25 + 0 * 0.25 + desc * 0.10
        # = 40 + 12.5 + 0 + ~5 = ~57.5 ... which is above 40 threshold
        # Need to use strict config to test this
        strict = MatchConfig(low_confidence_threshold=60.0, date_window_days=3)
        bank = [create_bank_txn(1, date(2026, 3, 1), 5000.00, "WIRE IN")]
        dms = [create_dms_txn(1, date(2026, 3, 5), 5000.00, "PAYMENT RECEIVED")]

        matches, _, _ = run_matching(bank, dms, strict)
        # Date score = 0 (4 days apart, window is 3)
        # Total = 100*0.40 + 50*0.25 + 0*0.25 + desc*0.10 ≈ 40+12.5+0+5 = 57.5
        # Below 60 threshold → should not be staged
        assert len(matches) == 0

    def test_match_ids_unique(self, default_config):
        """Each match gets a unique ID."""
        d = date(2026, 3, 1)
        bank = [
            create_bank_txn(1, d, -1000.00, "CHECK #100", check_number="100"),
            create_bank_txn(2, d, -2000.00, "CHECK #200", check_number="200"),
        ]
        dms = [
            create_dms_txn(1, d, -1000.00, "CHECK", reference="100", type_code="CHK"),
            create_dms_txn(2, d, -2000.00, "CHECK", reference="200", type_code="CHK"),
        ]

        matches, _, _ = run_matching(bank, dms, default_config)
        ids = [m.match_id for m in matches]
        assert len(ids) == len(set(ids))  # All unique

    def test_no_transaction_matched_twice(self, default_config):
        """No transaction appears in more than one match."""
        d = date(2026, 3, 1)
        bank = [
            create_bank_txn(1, d, 5000.00, "DEPOSIT"),
            create_bank_txn(2, d, 5000.00, "DEPOSIT"),  # Same amount
        ]
        dms = [
            create_dms_txn(1, d, 5000.00, "DEPOSIT A"),
        ]

        matches, unmatched_b, _ = run_matching(bank, dms, default_config)
        assert len(matches) == 1
        assert len(unmatched_b) == 1  # One bank txn should remain unmatched


class TestDuplicateAmountHandling:
    def test_two_same_amount_different_checks(self, default_config):
        """S06: Two identical amounts are correctly paired by check number."""
        d = date(2026, 3, 5)
        bank = [
            create_bank_txn(1, d, -2500.00, "CHECK #4530", check_number="4530"),
            create_bank_txn(2, d, -2500.00, "CHECK #4531", check_number="4531"),
        ]
        dms = [
            create_dms_txn(1, d, -2500.00, "PARTS SUPPLIER", reference="4530", type_code="CHK"),
            create_dms_txn(2, d, -2500.00, "UTILITY CO", reference="4531", type_code="CHK"),
        ]

        matches, _, _ = run_matching(bank, dms, default_config)
        assert len(matches) == 2

        # Verify correct pairing
        for m in matches:
            if 1 in m.bank_transaction_ids:
                assert 1 in m.dms_transaction_ids
            if 2 in m.bank_transaction_ids:
                assert 2 in m.dms_transaction_ids

    def test_two_same_amount_no_checks(self, default_config):
        """S07: Two identical amounts with no check numbers → both matched but order may vary."""
        d = date(2026, 3, 3)
        bank = [
            create_bank_txn(1, d, 4200.00, "DEPOSIT"),
            create_bank_txn(2, d, 4200.00, "DEPOSIT"),
        ]
        dms = [
            create_dms_txn(1, d, 4200.00, "CUSTOMER JONES"),
            create_dms_txn(2, d, 4200.00, "CUSTOMER WILLIAMS"),
        ]

        matches, _, _ = run_matching(bank, dms, default_config)
        assert len(matches) == 2
        # Both should be matched, but pairing is ambiguous (same score)
        matched_bank_ids = set()
        matched_dms_ids = set()
        for m in matches:
            matched_bank_ids.update(m.bank_transaction_ids)
            matched_dms_ids.update(m.dms_transaction_ids)
        assert matched_bank_ids == {1, 2}
        assert matched_dms_ids == {1, 2}

    def test_four_round_dollar_amounts(self, default_config):
        """S15: Four round-dollar transactions correctly paired by check numbers."""
        d = date(2026, 3, 7)
        bank = [
            create_bank_txn(1, d, -1000.00, "CHECK #4540", check_number="4540"),
            create_bank_txn(2, d, -1000.00, "CHECK #4541", check_number="4541"),
            create_bank_txn(3, d, -2000.00, "CHECK #4542", check_number="4542"),
            create_bank_txn(4, d, -2000.00, "CHECK #4543", check_number="4543"),
        ]
        dms = [
            create_dms_txn(1, d, -1000.00, "VENDOR A", reference="4540", type_code="CHK"),
            create_dms_txn(2, d, -1000.00, "VENDOR B", reference="4541", type_code="CHK"),
            create_dms_txn(3, d, -2000.00, "VENDOR C", reference="4542", type_code="CHK"),
            create_dms_txn(4, d, -2000.00, "VENDOR D", reference="4543", type_code="CHK"),
        ]

        matches, _, _ = run_matching(bank, dms, default_config)
        assert len(matches) == 4

        for m in matches:
            # Each bank should match its corresponding DMS (same check number)
            assert m.bank_transaction_ids == m.dms_transaction_ids
