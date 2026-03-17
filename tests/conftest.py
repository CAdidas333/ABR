"""
ABR Test Suite — Shared Fixtures and Utilities
"""

import csv
import json
import os
from datetime import date
from typing import Optional

import pytest

from matching_engine import (
    Transaction, MatchConfig, MatchResult,
    parse_bofa_csv, parse_truist_csv, parse_dms_csv,
)

DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')


@pytest.fixture
def default_config():
    """Default matching configuration."""
    return MatchConfig()


@pytest.fixture
def strict_config():
    """Strict configuration with higher thresholds."""
    return MatchConfig(
        high_confidence_threshold=90.0,
        medium_confidence_threshold=70.0,
        low_confidence_threshold=50.0,
    )


@pytest.fixture
def relaxed_config():
    """Relaxed configuration with lower thresholds."""
    return MatchConfig(
        high_confidence_threshold=75.0,
        medium_confidence_threshold=50.0,
        low_confidence_threshold=30.0,
        date_window_days=14,
    )


def create_bank_txn(
    txn_id: int,
    txn_date: date,
    amount: float,
    description: str = "DEPOSIT",
    check_number: str = "",
    bank_source: str = "BOFA",
) -> Transaction:
    """Create a bank transaction for testing."""
    return Transaction(
        transaction_id=txn_id,
        source="BANK",
        transaction_date=txn_date,
        description=description,
        amount=amount,
        check_number=check_number,
        bank_source=bank_source,
    )


def create_dms_txn(
    txn_id: int,
    txn_date: date,
    amount: float,
    description: str = "GL ENTRY",
    reference: str = "",
    type_code: str = "DEP",
    check_number: str = "",
) -> Transaction:
    """Create a DMS transaction for testing."""
    return Transaction(
        transaction_id=txn_id,
        source="DMS",
        transaction_date=txn_date,
        description=description,
        amount=amount,
        reference_number=reference,
        type_code=type_code,
        check_number=check_number,
    )


def load_scenario(scenario_id: str) -> dict:
    """Load a test scenario's expected results."""
    filepath = os.path.join(DATA_DIR, f'scenario_{scenario_id}_expected.json')
    with open(filepath, 'r') as f:
        return json.load(f)


def load_scenario_bank(scenario_id: str) -> list[Transaction]:
    """Load bank transactions from a scenario file."""
    filepath = os.path.join(DATA_DIR, f'scenario_{scenario_id}_bank.csv')
    expected = load_scenario(scenario_id)
    fmt = expected.get('bank_format', 'BOFA')
    if fmt == 'TRUIST':
        return parse_truist_csv(filepath)
    else:
        return parse_bofa_csv(filepath)


def load_scenario_dms(scenario_id: str) -> list[Transaction]:
    """Load DMS transactions from a scenario file."""
    filepath = os.path.join(DATA_DIR, f'scenario_{scenario_id}_dms.csv')
    return parse_dms_csv(filepath)
