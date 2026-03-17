"""
ABR Test Suite — CVR and Split Transaction Tests

Tests the many-to-one (CVR) and reverse split matching logic.
"""

from datetime import date, timedelta

import pytest

from matching_engine import (
    find_subset_sum, score_cvr_group, run_cvr_matching,
    run_reverse_split_matching, run_full_reconciliation,
    MatchConfig,
)
from conftest import create_bank_txn, create_dms_txn


# ---------------------------------------------------------------------------
# Subset Sum Solver
# ---------------------------------------------------------------------------

class TestSubsetSum:
    def test_two_fragments_exact(self):
        """S08: Two bank deposits sum to DMS amount exactly."""
        d = date(2026, 3, 1)
        candidates = [
            create_bank_txn(1, d, 5200.00),
            create_bank_txn(2, d, 4800.00),
        ]
        results = find_subset_sum(candidates, 10000.00, tolerance=0.01)
        assert len(results) >= 1
        assert any(len(r) == 2 for r in results)
        # Verify sum
        for result in results:
            assert abs(sum(t.amount for t in result) - 10000.00) <= 0.01

    def test_three_fragments_exact(self):
        """S09: Three bank deposits sum to DMS amount."""
        d = date(2026, 3, 2)
        candidates = [
            create_bank_txn(1, d, 5200.00),
            create_bank_txn(2, d, 4847.23),
            create_bank_txn(3, d, 5800.00),
        ]
        target = 15847.23
        results = find_subset_sum(candidates, target, tolerance=0.01)
        assert len(results) >= 1
        for result in results:
            assert abs(sum(t.amount for t in result) - target) <= 0.01

    def test_five_fragments_stress(self):
        """S10: Five fragments should still be found."""
        d = date(2026, 3, 3)
        fragments = [3200.00, 2150.50, 4500.00, 1875.25, 2274.25]
        candidates = [create_bank_txn(i + 1, d, amt) for i, amt in enumerate(fragments)]
        target = sum(fragments)  # 14000.00

        results = find_subset_sum(candidates, target, tolerance=0.01, max_depth=6)
        assert len(results) >= 1

    def test_ambiguous_multiple_solutions(self):
        """S11: Multiple valid subsets should all be returned."""
        d = date(2026, 3, 4)
        candidates = [
            create_bank_txn(1, d, 6000.00),
            create_bank_txn(2, d, 4000.00),
            create_bank_txn(3, d, 7000.00),
            create_bank_txn(4, d, 3000.00),
        ]
        target = 10000.00
        results = find_subset_sum(candidates, target, tolerance=0.01)
        # Should find at least 2 solutions: {6000,4000} and {7000,3000}
        assert len(results) >= 2

    def test_no_valid_combination(self):
        """No subset sums to target → empty results."""
        d = date(2026, 3, 1)
        candidates = [
            create_bank_txn(1, d, 1000.00),
            create_bank_txn(2, d, 2000.00),
        ]
        results = find_subset_sum(candidates, 5000.00, tolerance=0.01)
        assert len(results) == 0

    def test_tolerance_handling(self):
        """Sum within tolerance should be accepted."""
        d = date(2026, 3, 1)
        candidates = [
            create_bank_txn(1, d, 5000.005),
            create_bank_txn(2, d, 4999.995),  # Sum = 10000.00, target = 10000.00
        ]
        # With floats, use a wider tolerance to demonstrate the concept
        results = find_subset_sum(candidates, 10000.00, tolerance=0.02)
        assert len(results) >= 1

    def test_tolerance_exceeded(self):
        """Sum outside tolerance should not match."""
        d = date(2026, 3, 1)
        candidates = [
            create_bank_txn(1, d, 5000.00),
            create_bank_txn(2, d, 4999.00),  # Sum = 9999.00, off by 1.00
        ]
        results = find_subset_sum(candidates, 10000.00, tolerance=0.01)
        assert len(results) == 0

    def test_timeout_prevents_freeze(self):
        """Large candidate set should respect timeout."""
        d = date(2026, 3, 1)
        # 15 candidates with many possible combinations
        candidates = [create_bank_txn(i, d, 1000.00 + i * 100)
                      for i in range(1, 16)]
        target = 99999.00  # Unlikely to match

        import time
        start = time.time()
        results = find_subset_sum(candidates, target, timeout=0.5, max_depth=6)
        elapsed = time.time() - start

        # Should complete within reasonable time (timeout + overhead)
        assert elapsed < 2.0

    def test_single_candidate_not_matched(self):
        """Subsets must have at least 2 items."""
        d = date(2026, 3, 1)
        candidates = [create_bank_txn(1, d, 10000.00)]
        results = find_subset_sum(candidates, 10000.00, tolerance=0.01)
        assert len(results) == 0  # Min depth is 2


# ---------------------------------------------------------------------------
# CVR Confidence Scoring
# ---------------------------------------------------------------------------

class TestCVRConfidenceScoring:
    def test_exact_sum_high_confidence(self):
        """Exact sum, tight date cluster, 2 fragments → highest CVR confidence."""
        d = date(2026, 3, 1)
        group = [
            create_bank_txn(1, d, 5000.00),
            create_bank_txn(2, d, 5000.00),
        ]
        target = create_dms_txn(1, d, 10000.00, type_code="CVR")

        score = score_cvr_group(group, target)
        assert score >= 80.0

    def test_date_clustering_bonus(self):
        """Tighter date clustering → higher score."""
        d = date(2026, 3, 1)
        tight_group = [
            create_bank_txn(1, d, 5000.00),
            create_bank_txn(2, d, 5000.00),
        ]
        spread_group = [
            create_bank_txn(1, d, 5000.00),
            create_bank_txn(2, d + timedelta(days=6), 5000.00),
        ]
        target = create_dms_txn(1, d, 10000.00, type_code="CVR")

        tight_score = score_cvr_group(tight_group, target)
        spread_score = score_cvr_group(spread_group, target)

        assert tight_score > spread_score

    def test_many_fragments_lower_confidence(self):
        """More fragments → lower confidence."""
        d = date(2026, 3, 1)
        two_frag = [
            create_bank_txn(1, d, 5000.00),
            create_bank_txn(2, d, 5000.00),
        ]
        five_frag = [
            create_bank_txn(1, d, 2000.00),
            create_bank_txn(2, d, 2000.00),
            create_bank_txn(3, d, 2000.00),
            create_bank_txn(4, d, 2000.00),
            create_bank_txn(5, d, 2000.00),
        ]
        target = create_dms_txn(1, d, 10000.00, type_code="CVR")

        two_score = score_cvr_group(two_frag, target)
        five_score = score_cvr_group(five_frag, target)

        assert two_score > five_score


# ---------------------------------------------------------------------------
# CVR Matching Integration
# ---------------------------------------------------------------------------

class TestCVRDetection:
    def test_identifies_cvr_candidates(self):
        """CVR-type DMS entries should be found by the matcher."""
        d = date(2026, 3, 1)
        bank = [
            create_bank_txn(1, d, 5200.00),
            create_bank_txn(2, d, 4800.00),
        ]
        dms = [
            create_dms_txn(1, d, 10000.00, "CUSTOMER VEHICLE RECEIVABLE",
                          reference="CVR-001", type_code="CVR"),
        ]
        config = MatchConfig()
        matches = run_cvr_matching(bank, dms, config)
        assert len(matches) >= 1
        assert matches[0].match_type == "MANY_TO_ONE_BANK"

    def test_skips_already_matched_dms(self):
        """Already-matched DMS entries should be skipped."""
        d = date(2026, 3, 1)
        bank = [
            create_bank_txn(1, d, 5000.00),
            create_bank_txn(2, d, 5000.00),
        ]
        dms = [
            create_dms_txn(1, d, 10000.00, type_code="CVR"),
        ]
        dms[0].is_matched = True
        config = MatchConfig()
        matches = run_cvr_matching(bank, dms, config)
        assert len(matches) == 0


class TestReverseSplit:
    def test_one_bank_two_dms(self):
        """S12: Single bank deposit, two DMS entries that sum to it."""
        d = date(2026, 3, 5)
        bank = [create_bank_txn(1, d, 12400.00)]
        dms = [
            create_dms_txn(1, d, 7400.00, "PAYMENT FIRST HALF"),
            create_dms_txn(2, d, 5000.00, "PAYMENT SECOND HALF"),
        ]
        config = MatchConfig()
        matches = run_reverse_split_matching(bank, dms, config)
        assert len(matches) >= 1
        assert matches[0].match_type == "MANY_TO_ONE_DMS"


# ---------------------------------------------------------------------------
# Full Pipeline with CVR
# ---------------------------------------------------------------------------

class TestFullPipelineWithCVR:
    def test_mixed_one_to_one_and_cvr(self):
        """Pipeline handles both 1:1 and CVR matches in one pass."""
        d = date(2026, 3, 1)
        bank = [
            # 1:1 match
            create_bank_txn(1, d, -1250.00, "CHECK #4521", check_number="4521"),
            # CVR fragments
            create_bank_txn(2, d, 5200.00, "DEPOSIT"),
            create_bank_txn(3, d, 4800.00, "DEPOSIT"),
        ]
        dms = [
            # 1:1 match
            create_dms_txn(1, d, -1250.00, "CHECK PAYMENT", reference="4521",
                          type_code="CHK"),
            # CVR target
            create_dms_txn(2, d, 10000.00, "CVR - SMITH", reference="CVR-001",
                          type_code="CVR"),
        ]
        config = MatchConfig()
        result = run_full_reconciliation(bank, dms, config)

        assert len(result["one_to_one_matches"]) == 1
        assert len(result["cvr_matches"]) >= 1
        # Check 1:1 matched correctly
        oto = result["one_to_one_matches"][0]
        assert oto.confidence_score >= 95
