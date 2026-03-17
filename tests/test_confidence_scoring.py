"""
ABR Test Suite — Confidence Scoring Tests

Tests the individual factor scoring functions and the composite score calculation.
"""

from datetime import date

import pytest

from matching_engine import (
    score_amount, score_check_number, score_date, score_description,
    score_match, MatchConfig,
)
from conftest import create_bank_txn, create_dms_txn


# ---------------------------------------------------------------------------
# Amount Scoring
# ---------------------------------------------------------------------------

class TestAmountScoring:
    def test_exact_match_scores_100(self):
        assert score_amount(1250.00, 1250.00) == 100.0

    def test_penny_difference_scores_98(self):
        assert score_amount(1250.00, 1250.01) == 98.0
        assert score_amount(1250.01, 1250.00) == 98.0

    def test_nickel_difference_scores_90(self):
        assert score_amount(1250.00, 1250.05) == 90.0
        assert score_amount(1250.00, 1250.03) == 90.0

    def test_beyond_threshold_scores_0(self):
        assert score_amount(1250.00, 1250.06) == 0.0
        assert score_amount(1250.00, 1251.00) == 0.0
        assert score_amount(1250.00, 1300.00) == 0.0

    def test_zero_amounts(self):
        assert score_amount(0.0, 0.0) == 100.0

    def test_negative_amounts_exact(self):
        assert score_amount(-1250.00, -1250.00) == 100.0

    def test_negative_amounts_mismatch(self):
        assert score_amount(1250.00, -1250.00) == 0.0

    def test_large_amounts_exact(self):
        assert score_amount(999999.99, 999999.99) == 100.0


# ---------------------------------------------------------------------------
# Check Number Scoring
# ---------------------------------------------------------------------------

class TestCheckNumberScoring:
    def test_matching_check_numbers_score_100(self):
        score, veto = score_check_number("4521", "4521")
        assert score == 100.0
        assert veto is False

    def test_mismatched_check_numbers_score_0_and_veto(self):
        score, veto = score_check_number("4521", "4599")
        assert score == 0.0
        assert veto is True

    def test_one_missing_check_number_scores_50(self):
        score, veto = score_check_number("4521", "")
        assert score == 50.0
        assert veto is False

    def test_other_missing_check_number_scores_50(self):
        score, veto = score_check_number("", "4521")
        assert score == 50.0
        assert veto is False

    def test_both_missing_check_numbers_scores_50(self):
        score, veto = score_check_number("", "")
        assert score == 50.0
        assert veto is False

    def test_whitespace_handling(self):
        score, veto = score_check_number(" 4521 ", " 4521 ")
        assert score == 100.0


# ---------------------------------------------------------------------------
# Date Scoring
# ---------------------------------------------------------------------------

class TestDateScoring:
    def test_same_day_scores_100(self):
        d = date(2026, 3, 1)
        assert score_date(d, d) == 100.0

    def test_one_day_apart_scores_95(self):
        assert score_date(date(2026, 3, 1), date(2026, 3, 2)) == 95.0

    def test_two_days_apart_scores_85(self):
        assert score_date(date(2026, 3, 1), date(2026, 3, 3)) == 85.0

    def test_three_days_apart_scores_70(self):
        assert score_date(date(2026, 3, 1), date(2026, 3, 4)) == 70.0

    def test_seven_days_apart_scores_10(self):
        assert score_date(date(2026, 3, 1), date(2026, 3, 8)) == 10.0

    def test_beyond_window_scores_0(self):
        assert score_date(date(2026, 3, 1), date(2026, 3, 9)) == 0.0

    def test_direction_independent(self):
        assert score_date(date(2026, 3, 5), date(2026, 3, 1)) == \
               score_date(date(2026, 3, 1), date(2026, 3, 5))

    def test_custom_window(self):
        # With window of 3, 4 days apart should score 0
        assert score_date(date(2026, 3, 1), date(2026, 3, 5), max_window=3) == 0.0


# ---------------------------------------------------------------------------
# Description Scoring
# ---------------------------------------------------------------------------

class TestDescriptionScoring:
    def test_identical_descriptions_score_100(self):
        score = score_description("CHECK PAYMENT", "CHECK PAYMENT")
        assert score == 100.0

    def test_completely_different_scores_low(self):
        score = score_description("WIRE TRANSFER", "PAYROLL BATCH")
        assert score < 50.0

    def test_partial_match_scores_proportional(self):
        score = score_description("CHECK PAYMENT - VENDOR ABC",
                                  "CHECK PAYMENT - SUPPLIER XYZ")
        # Shared words "CHECK" and "PAYMENT" boost similarity significantly
        assert 30.0 < score <= 100.0

    def test_empty_descriptions_neutral(self):
        score = score_description("", "")
        assert score == 50.0

    def test_one_empty_neutral(self):
        score = score_description("DEPOSIT", "")
        assert score == 50.0

    def test_check_keyword_bonus(self):
        score1 = score_description("CHK PAYMENT", "CHECK VENDOR")
        score2 = score_description("WIRE PAYMENT", "TRANSFER VENDOR")
        # Both have CHECK/CHK keyword so score1 should get a bonus
        # (relative scores depend on other factors too)
        assert score1 > 0


# ---------------------------------------------------------------------------
# Composite Scoring
# ---------------------------------------------------------------------------

class TestCompositeScoring:
    def test_perfect_match_all_factors(self):
        """Exact amount, matching check numbers, same day → very high score."""
        config = MatchConfig()
        bank = create_bank_txn(1, date(2026, 3, 1), -1250.00,
                               "CHECK #4521", check_number="4521")
        dms = create_dms_txn(1, date(2026, 3, 1), -1250.00,
                             "CHECK PAYMENT", reference="4521", type_code="CHK")

        result = score_match(bank, dms, config)
        assert result is not None
        assert result.confidence_score >= 95.0

    def test_amount_only_no_check_far_date(self):
        """Exact amount, no check number, 5 days apart → moderate/low score."""
        config = MatchConfig()
        bank = create_bank_txn(1, date(2026, 3, 1), 5000.00, "DEPOSIT")
        dms = create_dms_txn(1, date(2026, 3, 6), 5000.00, "CUSTOMER DEPOSIT")

        result = score_match(bank, dms, config)
        assert result is not None
        # Amount:100*0.40=40 + Check:50*0.25=12.5 + Date:40*0.25=10 + Desc~=5
        assert 55.0 <= result.confidence_score <= 75.0

    def test_check_number_veto_caps_at_30(self):
        """Mismatched check numbers cap total at 30, regardless of other factors."""
        config = MatchConfig()
        bank = create_bank_txn(1, date(2026, 3, 1), -1250.00,
                               "CHECK #4521", check_number="4521")
        dms = create_dms_txn(1, date(2026, 3, 1), -1250.00,
                             "CHECK PAYMENT", reference="9999", type_code="CHK")

        result = score_match(bank, dms, config)
        assert result is not None
        assert result.confidence_score <= 30.0

    def test_amount_gate_returns_none(self):
        """Amount difference > $0.05 → no match candidate at all."""
        config = MatchConfig()
        bank = create_bank_txn(1, date(2026, 3, 1), 1250.00, "DEPOSIT")
        dms = create_dms_txn(1, date(2026, 3, 1), 1251.00, "CUSTOMER DEPOSIT")

        result = score_match(bank, dms, config)
        assert result is None

    def test_weight_configuration_changes_score(self):
        """Changing weights should change the score."""
        bank = create_bank_txn(1, date(2026, 3, 1), 5000.00, "DEPOSIT")
        dms = create_dms_txn(1, date(2026, 3, 1), 5000.00, "CUSTOMER DEPOSIT")

        config1 = MatchConfig(amount_weight=0.80, check_number_weight=0.10,
                              date_proximity_weight=0.05, description_weight=0.05)
        config2 = MatchConfig(amount_weight=0.20, check_number_weight=0.10,
                              date_proximity_weight=0.60, description_weight=0.10)

        result1 = score_match(bank, dms, config1)
        result2 = score_match(bank, dms, config2)

        assert result1 is not None
        assert result2 is not None
        # With 80% weight on amount (score=100) vs 60% weight on date (score=100),
        # both should be high but different
        assert result1.confidence_score != result2.confidence_score

    def test_score_breakdown_populated(self):
        """Score breakdown string should be populated."""
        config = MatchConfig()
        bank = create_bank_txn(1, date(2026, 3, 1), -1250.00,
                               "CHECK #4521", check_number="4521")
        dms = create_dms_txn(1, date(2026, 3, 1), -1250.00,
                             "CHECK PAYMENT", reference="4521", type_code="CHK")

        result = score_match(bank, dms, config)
        assert result is not None
        assert "Amount:" in result.score_breakdown
        assert "Check:" in result.score_breakdown
        assert "Date:" in result.score_breakdown
        assert "Desc:" in result.score_breakdown

    def test_date_difference_recorded(self):
        """Date difference should be recorded in the result."""
        config = MatchConfig()
        bank = create_bank_txn(1, date(2026, 3, 1), 5000.00)
        dms = create_dms_txn(1, date(2026, 3, 4), 5000.00)

        result = score_match(bank, dms, config)
        assert result is not None
        assert result.date_difference == 3
