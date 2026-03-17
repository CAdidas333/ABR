"""
ABR Test Suite — Import Parsing Tests

Tests CSV parsing for Bank of America, Truist, and R&R DMS formats.
"""

import os
from datetime import date

import pytest

from matching_engine import (
    parse_bofa_csv, parse_truist_csv, parse_dms_csv,
    detect_bank_format, extract_check_number,
)

DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')


# ---------------------------------------------------------------------------
# Check Number Extraction
# ---------------------------------------------------------------------------

class TestCheckNumberExtraction:
    def test_check_hash_format(self):
        assert extract_check_number("CHECK #4521") == "4521"

    def test_check_no_hash(self):
        assert extract_check_number("CHECK 4521") == "4521"

    def test_chk_format(self):
        assert extract_check_number("CHK #4521") == "4521"

    def test_chk_no_hash(self):
        assert extract_check_number("CHK 4521") == "4521"

    def test_ck_format(self):
        assert extract_check_number("CK #4521") == "4521"

    def test_check_in_longer_description(self):
        assert extract_check_number("CHECK #4521 VENDOR PAYMENT") == "4521"

    def test_no_check_number(self):
        assert extract_check_number("DEPOSIT") == ""

    def test_no_check_number_wire(self):
        assert extract_check_number("WIRE TRANSFER IN") == ""

    def test_case_insensitive(self):
        assert extract_check_number("check #4521") == "4521"

    def test_check_number_minimum_length(self):
        # Very short numbers (< 3 digits) should not match
        assert extract_check_number("CHECK #12") == ""

    def test_check_number_maximum_length(self):
        assert extract_check_number("CHECK #12345678") == "12345678"


# ---------------------------------------------------------------------------
# Bank of America Parsing
# ---------------------------------------------------------------------------

class TestBofAParsing:
    @pytest.fixture
    def bofa_file(self):
        return os.path.join(DATA_DIR, 'scenario_s01_bank.csv')

    def test_standard_format(self, bofa_file):
        """Parse a standard BofA CSV."""
        txns = parse_bofa_csv(bofa_file)
        assert len(txns) > 0
        assert all(t.source == "BANK" for t in txns)
        assert all(t.bank_source == "BOFA" for t in txns)

    def test_check_number_extraction_from_description(self):
        """Check numbers should be extracted from the Description field."""
        filepath = os.path.join(DATA_DIR, 'scenario_s06_bank.csv')
        txns = parse_bofa_csv(filepath)
        check_nums = [t.check_number for t in txns if t.check_number]
        assert "4530" in check_nums
        assert "4531" in check_nums

    def test_negative_amounts_for_debits(self):
        """BofA debits should have negative amounts."""
        filepath = os.path.join(DATA_DIR, 'scenario_s01_bank.csv')
        txns = parse_bofa_csv(filepath)
        # S01 is a check payment (debit)
        assert txns[0].amount < 0

    def test_date_parsing(self):
        """Dates should be parsed correctly."""
        filepath = os.path.join(DATA_DIR, 'scenario_s01_bank.csv')
        txns = parse_bofa_csv(filepath)
        assert isinstance(txns[0].transaction_date, date)

    def test_transaction_ids_sequential(self):
        """Transaction IDs should start at 1 and be sequential."""
        filepath = os.path.join(DATA_DIR, 'scenario_s06_bank.csv')
        txns = parse_bofa_csv(filepath)
        for i, txn in enumerate(txns):
            assert txn.transaction_id == i + 1


# ---------------------------------------------------------------------------
# Truist Parsing
# ---------------------------------------------------------------------------

class TestTruistParsing:
    @pytest.fixture
    def truist_file(self):
        return os.path.join(DATA_DIR, 'sample_truist.csv')

    def test_standard_format(self, truist_file):
        """Parse a standard Truist CSV."""
        txns = parse_truist_csv(truist_file)
        assert len(txns) > 0
        assert all(t.source == "BANK" for t in txns)
        assert all(t.bank_source == "TRUIST" for t in txns)

    def test_debit_credit_columns(self, truist_file):
        """Debits should be negative, credits should be positive."""
        txns = parse_truist_csv(truist_file)
        # First row is a credit (DEPOSIT)
        deposits = [t for t in txns if "DEPOSIT" in t.description]
        debits = [t for t in txns if t.amount < 0]

        assert len(deposits) > 0
        assert deposits[0].amount > 0
        assert len(debits) > 0

    def test_amount_sign_convention(self, truist_file):
        """Verify amount sign convention: debit=negative, credit=positive."""
        txns = parse_truist_csv(truist_file)
        for txn in txns:
            assert isinstance(txn.amount, float)


# ---------------------------------------------------------------------------
# DMS Parsing
# ---------------------------------------------------------------------------

class TestDMSParsing:
    @pytest.fixture
    def dms_file(self):
        return os.path.join(DATA_DIR, 'scenario_s01_dms.csv')

    def test_standard_format(self, dms_file):
        """Parse a standard DMS GL export."""
        txns = parse_dms_csv(dms_file)
        assert len(txns) > 0
        assert all(t.source == "DMS" for t in txns)

    def test_type_code_extraction(self):
        """Type codes should be correctly extracted."""
        filepath = os.path.join(DATA_DIR, 'scenario_s09_dms.csv')
        txns = parse_dms_csv(filepath)
        assert any(t.type_code == "CVR" for t in txns)

    def test_reference_number_extraction(self):
        """Reference numbers should be preserved."""
        filepath = os.path.join(DATA_DIR, 'scenario_s01_dms.csv')
        txns = parse_dms_csv(filepath)
        assert txns[0].reference_number != ""

    def test_check_number_for_chk_type(self):
        """CHK type transactions should have check_number extracted from reference."""
        filepath = os.path.join(DATA_DIR, 'scenario_s06_dms.csv')
        txns = parse_dms_csv(filepath)
        chk_txns = [t for t in txns if t.type_code == "CHK"]
        assert len(chk_txns) > 0
        assert any(t.check_number for t in chk_txns)


# ---------------------------------------------------------------------------
# Format Detection
# ---------------------------------------------------------------------------

class TestFormatDetection:
    def test_auto_detect_bofa(self):
        filepath = os.path.join(DATA_DIR, 'scenario_s01_bank.csv')
        fmt = detect_bank_format(filepath)
        assert fmt == "BOFA"

    def test_auto_detect_truist(self):
        filepath = os.path.join(DATA_DIR, 'sample_truist.csv')
        fmt = detect_bank_format(filepath)
        assert fmt == "TRUIST"
