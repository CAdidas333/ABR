#!/usr/bin/env python3
"""
ABR Test Data Generator

Generates realistic bank statement CSVs (BofA + Truist), R&R DMS GL exports,
and outstanding items files for testing the matching algorithm.

Covers 21 test scenarios including exact matches, duplicates, CVR fragments,
splits, outstanding carry-forward, and a full-month simulation.
"""

import csv
import json
import os
import random
from datetime import date, timedelta

# Fixed seed for reproducibility
random.seed(42)

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), '..', 'tests', 'data')


def ensure_dir():
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def write_bofa_csv(filename: str, rows: list[dict]):
    """Write Bank of America format CSV."""
    filepath = os.path.join(OUTPUT_DIR, filename)
    with open(filepath, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['Date', 'Description', 'Amount', 'Running Balance'])
        writer.writeheader()
        writer.writerows(rows)


def write_truist_csv(filename: str, rows: list[dict]):
    """Write Truist format CSV."""
    filepath = os.path.join(OUTPUT_DIR, filename)
    with open(filepath, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['Date', 'Description', 'Debit', 'Credit', 'Balance'])
        writer.writeheader()
        writer.writerows(rows)


def write_dms_csv(filename: str, rows: list[dict]):
    """Write R&R DMS GL export format CSV."""
    filepath = os.path.join(OUTPUT_DIR, filename)
    with open(filepath, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['GL Date', 'Description', 'Reference', 'Amount', 'Type Code'])
        writer.writeheader()
        writer.writerows(rows)


def write_scenario(scenario_id: str, description: str, bank_rows: list[dict],
                   dms_rows: list[dict], expected: dict,
                   bank_format: str = "BOFA"):
    """Write a complete test scenario with bank data, DMS data, and expected results."""
    if bank_format == "BOFA":
        write_bofa_csv(f'scenario_{scenario_id}_bank.csv', bank_rows)
    else:
        write_truist_csv(f'scenario_{scenario_id}_bank.csv', bank_rows)

    write_dms_csv(f'scenario_{scenario_id}_dms.csv', dms_rows)

    expected_data = {
        'scenario_id': scenario_id,
        'description': description,
        'bank_format': bank_format,
        'expected': expected,
    }
    filepath = os.path.join(OUTPUT_DIR, f'scenario_{scenario_id}_expected.json')
    with open(filepath, 'w') as f:
        json.dump(expected_data, f, indent=2, default=str)


def fmt_date(d: date) -> str:
    return d.strftime('%m/%d/%Y')


# ---------------------------------------------------------------------------
# Scenario generators
# ---------------------------------------------------------------------------

def generate_s01():
    """S01: Perfect 1:1 match — check with number, same day."""
    d = date(2026, 3, 2)
    bank = [{'Date': fmt_date(d), 'Description': 'CHECK #4521',
             'Amount': '-1250.00', 'Running Balance': '98750.00'}]
    dms = [{'GL Date': fmt_date(d), 'Description': 'CHECK PAYMENT - VENDOR ABC',
            'Reference': '4521', 'Amount': '-1250.00', 'Type Code': 'CHK'}]
    expected = {
        'match_count': 1,
        'matches': [{'bank_idx': 0, 'dms_idx': 0, 'type': 'ONE_TO_ONE',
                      'confidence_min': 95, 'confidence_max': 100}],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s01', 'Perfect 1:1 match - check with number, same day', bank, dms, expected)


def generate_s02():
    """S02: Perfect 1:1 match — deposit, same day, no check number."""
    d = date(2026, 3, 1)
    bank = [{'Date': fmt_date(d), 'Description': 'DEPOSIT',
             'Amount': '8500.00', 'Running Balance': '108500.00'}]
    dms = [{'GL Date': fmt_date(d), 'Description': 'CUSTOMER DEPOSIT',
            'Reference': 'DEP-0301', 'Amount': '8500.00', 'Type Code': 'DEP'}]
    expected = {
        'match_count': 1,
        'matches': [{'bank_idx': 0, 'dms_idx': 0, 'type': 'ONE_TO_ONE',
                      'confidence_min': 70, 'confidence_max': 85}],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s02', 'Perfect 1:1 match - deposit, same day, no check #', bank, dms, expected)


def generate_s03():
    """S03: Amount match but wrong check number — should be vetoed."""
    d = date(2026, 3, 2)
    bank = [{'Date': fmt_date(d), 'Description': 'CHECK #4521',
             'Amount': '-1250.00', 'Running Balance': '98750.00'}]
    dms = [{'GL Date': fmt_date(d), 'Description': 'CHECK PAYMENT - VENDOR XYZ',
            'Reference': '4599', 'Amount': '-1250.00', 'Type Code': 'CHK'}]
    expected = {
        'match_count': 1,
        'matches': [{'bank_idx': 0, 'dms_idx': 0, 'type': 'ONE_TO_ONE',
                      'confidence_min': 0, 'confidence_max': 30}],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s03', 'Amount match but wrong check number - vetoed', bank, dms, expected)


def generate_s04():
    """S04: Amount match, date 3 days apart."""
    d_bank = date(2026, 3, 1)
    d_dms = date(2026, 3, 4)
    bank = [{'Date': fmt_date(d_bank), 'Description': 'ACH CREDIT - INSURANCE REFUND',
             'Amount': '3200.00', 'Running Balance': '103200.00'}]
    dms = [{'GL Date': fmt_date(d_dms), 'Description': 'INSURANCE REFUND RECEIVED',
            'Reference': 'ACH-0304', 'Amount': '3200.00', 'Type Code': 'ACH'}]
    expected = {
        'match_count': 1,
        'matches': [{'bank_idx': 0, 'dms_idx': 0, 'type': 'ONE_TO_ONE',
                      'confidence_min': 60, 'confidence_max': 80}],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s04', 'Amount match, date 3 days apart', bank, dms, expected)


def generate_s05():
    """S05: Amount match, date 7 days apart — low confidence."""
    d_bank = date(2026, 3, 1)
    d_dms = date(2026, 3, 8)
    bank = [{'Date': fmt_date(d_bank), 'Description': 'WIRE TRANSFER IN',
             'Amount': '15000.00', 'Running Balance': '115000.00'}]
    dms = [{'GL Date': fmt_date(d_dms), 'Description': 'WIRE - MANUFACTURER REBATE',
            'Reference': 'WIR-0308', 'Amount': '15000.00', 'Type Code': 'WIR'}]
    expected = {
        'match_count': 1,
        'matches': [{'bank_idx': 0, 'dms_idx': 0, 'type': 'ONE_TO_ONE',
                      'confidence_min': 40, 'confidence_max': 60}],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s05', 'Amount match, date 7 days apart - low confidence', bank, dms, expected)


def generate_s06():
    """S06: Duplicate dollar amount, differentiated by check number."""
    d = date(2026, 3, 5)
    bank = [
        {'Date': fmt_date(d), 'Description': 'CHECK #4530',
         'Amount': '-2500.00', 'Running Balance': '97500.00'},
        {'Date': fmt_date(d), 'Description': 'CHECK #4531',
         'Amount': '-2500.00', 'Running Balance': '95000.00'},
    ]
    dms = [
        {'GL Date': fmt_date(d), 'Description': 'CHECK PAYMENT - PARTS SUPPLIER',
         'Reference': '4530', 'Amount': '-2500.00', 'Type Code': 'CHK'},
        {'GL Date': fmt_date(d), 'Description': 'CHECK PAYMENT - UTILITY CO',
         'Reference': '4531', 'Amount': '-2500.00', 'Type Code': 'CHK'},
    ]
    expected = {
        'match_count': 2,
        'matches': [
            {'bank_idx': 0, 'dms_idx': 0, 'type': 'ONE_TO_ONE',
             'confidence_min': 95, 'confidence_max': 100},
            {'bank_idx': 1, 'dms_idx': 1, 'type': 'ONE_TO_ONE',
             'confidence_min': 95, 'confidence_max': 100},
        ],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s06', 'Duplicate dollar amount, differentiated by check #', bank, dms, expected)


def generate_s07():
    """S07: Duplicate dollar amount, no check numbers — both flagged."""
    d = date(2026, 3, 3)
    bank = [
        {'Date': fmt_date(d), 'Description': 'DEPOSIT',
         'Amount': '4200.00', 'Running Balance': '108200.00'},
        {'Date': fmt_date(d), 'Description': 'DEPOSIT',
         'Amount': '4200.00', 'Running Balance': '112400.00'},
    ]
    dms = [
        {'GL Date': fmt_date(d), 'Description': 'CUSTOMER PAYMENT - JONES',
         'Reference': 'DEP-0303A', 'Amount': '4200.00', 'Type Code': 'DEP'},
        {'GL Date': fmt_date(d), 'Description': 'CUSTOMER PAYMENT - WILLIAMS',
         'Reference': 'DEP-0303B', 'Amount': '4200.00', 'Type Code': 'DEP'},
    ]
    expected = {
        'match_count': 2,
        'matches': [
            {'bank_idx': 0, 'dms_idx': 0, 'type': 'ONE_TO_ONE',
             'confidence_min': 70, 'confidence_max': 85},
            {'bank_idx': 1, 'dms_idx': 1, 'type': 'ONE_TO_ONE',
             'confidence_min': 70, 'confidence_max': 85},
        ],
        'note': 'Both should be flagged as duplicate-amount matches for careful review',
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s07', 'Duplicate dollar amount, no check numbers', bank, dms, expected)


def generate_s08():
    """S08: CVR 2-fragment — bank splits DMS entry."""
    d = date(2026, 3, 1)
    bank = [
        {'Date': fmt_date(d), 'Description': 'DEPOSIT',
         'Amount': '5200.00', 'Running Balance': '105200.00'},
        {'Date': fmt_date(d + timedelta(days=1)), 'Description': 'DEPOSIT',
         'Amount': '4800.00', 'Running Balance': '110000.00'},
    ]
    dms = [
        {'GL Date': fmt_date(d), 'Description': 'CUSTOMER VEHICLE RECEIVABLE - SMITH',
         'Reference': 'CVR-2026-0301', 'Amount': '10000.00', 'Type Code': 'CVR'},
    ]
    expected = {
        'match_count': 0,
        'one_to_one_matches': 0,
        'cvr_matches': 1,
        'cvr_groups': [{'bank_indices': [0, 1], 'dms_idx': 0, 'sum': 10000.00,
                        'confidence_min': 75}],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s08', 'CVR 2-fragment - bank splits DMS entry', bank, dms, expected)


def generate_s09():
    """S09: CVR 3-fragment."""
    d = date(2026, 3, 2)
    bank = [
        {'Date': fmt_date(d), 'Description': 'DEPOSIT',
         'Amount': '5200.00', 'Running Balance': '105200.00'},
        {'Date': fmt_date(d), 'Description': 'DEPOSIT',
         'Amount': '4847.23', 'Running Balance': '110047.23'},
        {'Date': fmt_date(d + timedelta(days=1)), 'Description': 'DEPOSIT',
         'Amount': '5800.00', 'Running Balance': '115847.23'},
    ]
    dms = [
        {'GL Date': fmt_date(d), 'Description': 'CUSTOMER VEHICLE RECEIVABLE - DOE',
         'Reference': 'CVR-2026-0302', 'Amount': '15847.23', 'Type Code': 'CVR'},
    ]
    expected = {
        'match_count': 0,
        'one_to_one_matches': 0,
        'cvr_matches': 1,
        'cvr_groups': [{'bank_indices': [0, 1, 2], 'dms_idx': 0, 'sum': 15847.23,
                        'confidence_min': 70}],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s09', 'CVR 3-fragment', bank, dms, expected)


def generate_s10():
    """S10: CVR 5-fragment stress test."""
    d = date(2026, 3, 3)
    fragments = [3200.00, 2150.50, 4500.00, 1875.25, 2274.25]
    total = sum(fragments)  # 14000.00

    bank = []
    balance = 100000.00
    for i, amt in enumerate(fragments):
        balance += amt
        bank.append({
            'Date': fmt_date(d + timedelta(days=i % 3)),
            'Description': 'DEPOSIT',
            'Amount': f'{amt:.2f}',
            'Running Balance': f'{balance:.2f}'
        })

    dms = [
        {'GL Date': fmt_date(d), 'Description': 'CUSTOMER VEHICLE RECEIVABLE - GARCIA',
         'Reference': 'CVR-2026-0303', 'Amount': f'{total:.2f}', 'Type Code': 'CVR'},
    ]
    expected = {
        'match_count': 0,
        'one_to_one_matches': 0,
        'cvr_matches': 1,
        'cvr_groups': [{'fragment_count': 5, 'dms_idx': 0, 'sum': total,
                        'confidence_min': 55}],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s10', 'CVR 5-fragment stress test', bank, dms, expected)


def generate_s11():
    """S11: CVR ambiguous — two possible groupings."""
    d = date(2026, 3, 4)
    # Target: 10000.00
    # Group A: 6000 + 4000 = 10000
    # Group B: 7000 + 3000 = 10000
    bank = [
        {'Date': fmt_date(d), 'Description': 'DEPOSIT',
         'Amount': '6000.00', 'Running Balance': '106000.00'},
        {'Date': fmt_date(d), 'Description': 'DEPOSIT',
         'Amount': '4000.00', 'Running Balance': '110000.00'},
        {'Date': fmt_date(d + timedelta(days=1)), 'Description': 'DEPOSIT',
         'Amount': '7000.00', 'Running Balance': '117000.00'},
        {'Date': fmt_date(d + timedelta(days=1)), 'Description': 'DEPOSIT',
         'Amount': '3000.00', 'Running Balance': '120000.00'},
    ]
    dms = [
        {'GL Date': fmt_date(d), 'Description': 'CUSTOMER VEHICLE RECEIVABLE - PATEL',
         'Reference': 'CVR-2026-0304', 'Amount': '10000.00', 'Type Code': 'CVR'},
    ]
    expected = {
        'match_count': 0,
        'one_to_one_matches': 0,
        'cvr_matches_min': 2,
        'note': 'Multiple valid groupings should be surfaced for controller review',
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s11', 'CVR ambiguous - two possible groupings', bank, dms, expected)


def generate_s12():
    """S12: Reverse split — 1 bank deposit, 2 DMS entries."""
    d = date(2026, 3, 5)
    bank = [
        {'Date': fmt_date(d), 'Description': 'DEPOSIT',
         'Amount': '12400.00', 'Running Balance': '112400.00'},
    ]
    dms = [
        {'GL Date': fmt_date(d), 'Description': 'CUSTOMER PAYMENT - FIRST HALF',
         'Reference': 'DEP-0305A', 'Amount': '7400.00', 'Type Code': 'DEP'},
        {'GL Date': fmt_date(d), 'Description': 'CUSTOMER PAYMENT - SECOND HALF',
         'Reference': 'DEP-0305B', 'Amount': '5000.00', 'Type Code': 'DEP'},
    ]
    expected = {
        'match_count': 0,
        'one_to_one_matches': 0,
        'split_matches': 1,
        'split_groups': [{'bank_idx': 0, 'dms_indices': [0, 1], 'sum': 12400.00}],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s12', 'Reverse split - 1 bank, 2 DMS', bank, dms, expected)


def generate_s13():
    """S13: Outstanding item from prior period matches new data."""
    d = date(2026, 3, 10)
    bank = [
        {'Date': fmt_date(d), 'Description': 'CHECK #4450',
         'Amount': '-890.00', 'Running Balance': '99110.00'},
    ]
    dms = [
        {'GL Date': fmt_date(date(2026, 2, 25)), 'Description': 'CHECK PAYMENT - OFFICE SUPPLIES',
         'Reference': '4450', 'Amount': '-890.00', 'Type Code': 'CHK'},
    ]
    expected = {
        'match_count': 1,
        'matches': [{'bank_idx': 0, 'dms_idx': 0, 'type': 'ONE_TO_ONE',
                      'confidence_min': 55, 'confidence_max': 80,
                      'note': 'Date 13 days apart but check numbers match'}],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    # Note: with default 7-day window, this won't match by date.
    # But check number match should still produce a score.
    # We need to extend the date window for outstanding items.
    write_scenario('s13', 'Outstanding item from prior period', bank, dms, expected)


def generate_s14():
    """S14: Void / negative amount check."""
    d = date(2026, 3, 6)
    bank = [
        {'Date': fmt_date(d), 'Description': 'CHECK #4535 VOID',
         'Amount': '1500.00', 'Running Balance': '101500.00'},
    ]
    dms = [
        {'GL Date': fmt_date(d), 'Description': 'VOID CHECK - VENDOR RETURN',
         'Reference': '4535', 'Amount': '1500.00', 'Type Code': 'CHK'},
    ]
    expected = {
        'match_count': 1,
        'matches': [{'bank_idx': 0, 'dms_idx': 0, 'type': 'ONE_TO_ONE',
                      'confidence_min': 90, 'confidence_max': 100}],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s14', 'Void / negative amount check', bank, dms, expected)


def generate_s15():
    """S15: Round dollar amounts ($1000, $2000) — multiple duplicates."""
    d = date(2026, 3, 7)
    bank = [
        {'Date': fmt_date(d), 'Description': 'CHECK #4540',
         'Amount': '-1000.00', 'Running Balance': '99000.00'},
        {'Date': fmt_date(d), 'Description': 'CHECK #4541',
         'Amount': '-1000.00', 'Running Balance': '98000.00'},
        {'Date': fmt_date(d + timedelta(days=1)), 'Description': 'CHECK #4542',
         'Amount': '-2000.00', 'Running Balance': '96000.00'},
        {'Date': fmt_date(d + timedelta(days=1)), 'Description': 'CHECK #4543',
         'Amount': '-2000.00', 'Running Balance': '94000.00'},
    ]
    dms = [
        {'GL Date': fmt_date(d), 'Description': 'CHECK - VENDOR A',
         'Reference': '4540', 'Amount': '-1000.00', 'Type Code': 'CHK'},
        {'GL Date': fmt_date(d), 'Description': 'CHECK - VENDOR B',
         'Reference': '4541', 'Amount': '-1000.00', 'Type Code': 'CHK'},
        {'GL Date': fmt_date(d + timedelta(days=1)), 'Description': 'CHECK - VENDOR C',
         'Reference': '4542', 'Amount': '-2000.00', 'Type Code': 'CHK'},
        {'GL Date': fmt_date(d + timedelta(days=1)), 'Description': 'CHECK - VENDOR D',
         'Reference': '4543', 'Amount': '-2000.00', 'Type Code': 'CHK'},
    ]
    expected = {
        'match_count': 4,
        'matches': [
            {'bank_idx': 0, 'dms_idx': 0, 'confidence_min': 95},
            {'bank_idx': 1, 'dms_idx': 1, 'confidence_min': 95},
            {'bank_idx': 2, 'dms_idx': 2, 'confidence_min': 95},
            {'bank_idx': 3, 'dms_idx': 3, 'confidence_min': 95},
        ],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s15', 'Round dollar amounts with multiple duplicates', bank, dms, expected)


def generate_s16():
    """S16: ACH payroll — large exact match, same day."""
    d = date(2026, 3, 2)
    bank = [{'Date': fmt_date(d), 'Description': 'ACH DEBIT - ADP PAYROLL',
             'Amount': '-45678.90', 'Running Balance': '54321.10'}]
    dms = [{'GL Date': fmt_date(d), 'Description': 'PAYROLL BATCH 2026-03',
            'Reference': 'PAY-0302', 'Amount': '-45678.90', 'Type Code': 'PAY'}]
    expected = {
        'match_count': 1,
        'matches': [{'bank_idx': 0, 'dms_idx': 0, 'type': 'ONE_TO_ONE',
                      'confidence_min': 75, 'confidence_max': 85}],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s16', 'ACH payroll - large exact match, same day', bank, dms, expected)


def generate_s17():
    """S17: Wire transfer with reference number match."""
    d = date(2026, 3, 3)
    bank = [{'Date': fmt_date(d), 'Description': 'WIRE TRANSFER IN REF WIR-0303',
             'Amount': '25000.00', 'Running Balance': '125000.00'}]
    dms = [{'GL Date': fmt_date(d), 'Description': 'WIRE - MANUFACTURER ALLOCATION',
            'Reference': 'WIR-0303', 'Amount': '25000.00', 'Type Code': 'WIR'}]
    expected = {
        'match_count': 1,
        'matches': [{'bank_idx': 0, 'dms_idx': 0, 'type': 'ONE_TO_ONE',
                      'confidence_min': 70, 'confidence_max': 85}],
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s17', 'Wire transfer with reference number match', bank, dms, expected)


def generate_s18():
    """S18: Bank fee — small recurring amount, multiple months."""
    bank = [
        {'Date': fmt_date(date(2026, 3, 5)), 'Description': 'BANK FEE - SERVICE CHARGE',
         'Amount': '-45.00', 'Running Balance': '99955.00'},
        {'Date': fmt_date(date(2026, 3, 15)), 'Description': 'BANK FEE - SERVICE CHARGE',
         'Amount': '-45.00', 'Running Balance': '99910.00'},
        {'Date': fmt_date(date(2026, 3, 25)), 'Description': 'BANK FEE - SERVICE CHARGE',
         'Amount': '-45.00', 'Running Balance': '99865.00'},
    ]
    dms = [
        {'GL Date': fmt_date(date(2026, 3, 5)), 'Description': 'BANK SERVICE FEE',
         'Reference': 'FEE-0305', 'Amount': '-45.00', 'Type Code': 'FEE'},
        {'GL Date': fmt_date(date(2026, 3, 15)), 'Description': 'BANK SERVICE FEE',
         'Reference': 'FEE-0315', 'Amount': '-45.00', 'Type Code': 'FEE'},
        {'GL Date': fmt_date(date(2026, 3, 25)), 'Description': 'BANK SERVICE FEE',
         'Reference': 'FEE-0325', 'Amount': '-45.00', 'Type Code': 'FEE'},
    ]
    expected = {
        'match_count': 3,
        'matches': [
            {'bank_idx': 0, 'dms_idx': 0, 'confidence_min': 70},
            {'bank_idx': 1, 'dms_idx': 1, 'confidence_min': 70},
            {'bank_idx': 2, 'dms_idx': 2, 'confidence_min': 70},
        ],
        'note': 'Date proximity should differentiate same-amount recurring fees',
        'unmatched_bank': 0, 'unmatched_dms': 0
    }
    write_scenario('s18', 'Bank fee - recurring same amount', bank, dms, expected)


def generate_s19():
    """S19: Orphan bank transaction — no DMS match exists."""
    d = date(2026, 3, 8)
    bank = [{'Date': fmt_date(d), 'Description': 'MISC CREDIT',
             'Amount': '333.33', 'Running Balance': '100333.33'}]
    dms = []
    expected = {
        'match_count': 0,
        'unmatched_bank': 1, 'unmatched_dms': 0
    }
    write_scenario('s19', 'Orphan bank transaction - no DMS match', bank, dms, expected)


def generate_s20():
    """S20: Orphan DMS transaction — no bank match exists."""
    d = date(2026, 3, 9)
    bank = []
    dms = [{'GL Date': fmt_date(d), 'Description': 'ADJUSTMENT - JOURNAL ENTRY',
            'Reference': 'ADJ-0309', 'Amount': '555.55', 'Type Code': 'ADJ'}]
    expected = {
        'match_count': 0,
        'unmatched_bank': 0, 'unmatched_dms': 1
    }
    write_scenario('s20', 'Orphan DMS transaction - no bank match', bank, dms, expected)


def generate_s21():
    """S21: Full month simulation — ~150 bank + ~120 DMS transactions."""
    random.seed(42)
    base_date = date(2026, 3, 1)
    bank_rows = []
    dms_rows = []
    balance = 250000.00

    # Track expected matches for validation
    match_pairs = []

    # --- Generate matched pairs first ---

    # 40 check payments (exact match with check numbers)
    for i in range(40):
        check_num = 4500 + i
        day_offset = random.randint(0, 27)
        d = base_date + timedelta(days=day_offset)
        dms_day_offset = random.randint(0, 2)  # DMS within 0-2 days
        d_dms = d + timedelta(days=dms_day_offset)
        amount = round(random.uniform(200, 15000), 2)

        balance -= amount
        bank_rows.append({
            'Date': fmt_date(d),
            'Description': f'CHECK #{check_num}',
            'Amount': f'{-amount:.2f}',
            'Running Balance': f'{balance:.2f}'
        })
        dms_rows.append({
            'GL Date': fmt_date(d_dms),
            'Description': f'CHECK PAYMENT - VENDOR {chr(65 + i % 26)}{i // 26}',
            'Reference': str(check_num),
            'Amount': f'{-amount:.2f}',
            'Type Code': 'CHK'
        })
        match_pairs.append(('CHECK', i))

    # 25 deposits (amount + date match, no check number)
    for i in range(25):
        day_offset = random.randint(0, 27)
        d = base_date + timedelta(days=day_offset)
        dms_day_offset = random.randint(0, 1)
        d_dms = d + timedelta(days=dms_day_offset)
        amount = round(random.uniform(1000, 50000), 2)

        balance += amount
        bank_rows.append({
            'Date': fmt_date(d),
            'Description': 'DEPOSIT',
            'Amount': f'{amount:.2f}',
            'Running Balance': f'{balance:.2f}'
        })
        dms_rows.append({
            'GL Date': fmt_date(d_dms),
            'Description': f'CUSTOMER DEPOSIT - CUST {i + 1}',
            'Reference': f'DEP-{d_dms.strftime("%m%d")}{chr(65 + i)}',
            'Amount': f'{amount:.2f}',
            'Type Code': 'DEP'
        })
        match_pairs.append(('DEPOSIT', i))

    # 10 ACH transactions
    for i in range(10):
        day_offset = random.randint(0, 27)
        d = base_date + timedelta(days=day_offset)
        amount = round(random.uniform(500, 25000), 2)

        balance -= amount
        bank_rows.append({
            'Date': fmt_date(d),
            'Description': f'ACH DEBIT - VENDOR {i + 1}',
            'Amount': f'{-amount:.2f}',
            'Running Balance': f'{balance:.2f}'
        })
        dms_rows.append({
            'GL Date': fmt_date(d),
            'Description': f'ACH PAYMENT - VENDOR {i + 1}',
            'Reference': f'ACH-{d.strftime("%m%d")}{i}',
            'Amount': f'{-amount:.2f}',
            'Type Code': 'ACH'
        })
        match_pairs.append(('ACH', i))

    # 5 wire transfers
    for i in range(5):
        day_offset = random.randint(0, 27)
        d = base_date + timedelta(days=day_offset)
        amount = round(random.uniform(10000, 100000), 2)

        balance += amount
        bank_rows.append({
            'Date': fmt_date(d),
            'Description': f'WIRE TRANSFER IN',
            'Amount': f'{amount:.2f}',
            'Running Balance': f'{balance:.2f}'
        })
        dms_rows.append({
            'GL Date': fmt_date(d + timedelta(days=random.randint(0, 1))),
            'Description': f'WIRE - MANUFACTURER {i + 1}',
            'Reference': f'WIR-{d.strftime("%m%d")}{i}',
            'Amount': f'{amount:.2f}',
            'Type Code': 'WIR'
        })
        match_pairs.append(('WIRE', i))

    # 3 CVR many-to-one scenarios
    for i in range(3):
        day_offset = random.randint(0, 20)
        d = base_date + timedelta(days=day_offset)
        num_fragments = random.randint(2, 4)
        fragments = [round(random.uniform(2000, 8000), 2) for _ in range(num_fragments)]
        total = sum(fragments)

        for j, frag in enumerate(fragments):
            balance += frag
            bank_rows.append({
                'Date': fmt_date(d + timedelta(days=j % 2)),
                'Description': 'DEPOSIT',
                'Amount': f'{frag:.2f}',
                'Running Balance': f'{balance:.2f}'
            })

        dms_rows.append({
            'GL Date': fmt_date(d),
            'Description': f'CUSTOMER VEHICLE RECEIVABLE - CVR{i + 1}',
            'Reference': f'CVR-2026-{d.strftime("%m%d")}{i}',
            'Amount': f'{total:.2f}',
            'Type Code': 'CVR'
        })
        match_pairs.append(('CVR', i, num_fragments))

    # 15 additional bank transactions that should match with moderate confidence
    for i in range(15):
        day_offset = random.randint(0, 27)
        d = base_date + timedelta(days=day_offset)
        d_dms = d + timedelta(days=random.randint(2, 5))
        amount = round(random.uniform(100, 5000), 2)

        balance -= amount
        bank_rows.append({
            'Date': fmt_date(d),
            'Description': f'MISC PAYMENT {i + 1}',
            'Amount': f'{-amount:.2f}',
            'Running Balance': f'{balance:.2f}'
        })
        dms_rows.append({
            'GL Date': fmt_date(d_dms),
            'Description': f'MISCELLANEOUS EXPENSE {i + 1}',
            'Reference': f'MISC-{d_dms.strftime("%m%d")}{i}',
            'Amount': f'{-amount:.2f}',
            'Type Code': 'OTH'
        })

    # --- Orphan transactions (no match) ---

    # 15 orphan bank transactions
    for i in range(15):
        day_offset = random.randint(0, 27)
        d = base_date + timedelta(days=day_offset)
        amount = round(random.uniform(50, 3000), 2)
        if random.random() > 0.5:
            amount = -amount
        balance += amount
        bank_rows.append({
            'Date': fmt_date(d),
            'Description': f'UNMATCHED BANK TXN {i + 1}',
            'Amount': f'{amount:.2f}',
            'Running Balance': f'{balance:.2f}'
        })

    # 10 orphan DMS transactions
    for i in range(10):
        day_offset = random.randint(0, 27)
        d = base_date + timedelta(days=day_offset)
        amount = round(random.uniform(50, 3000), 2)
        if random.random() > 0.5:
            amount = -amount
        dms_rows.append({
            'GL Date': fmt_date(d),
            'Description': f'UNMATCHED DMS TXN {i + 1}',
            'Reference': f'UNM-{d.strftime("%m%d")}{i}',
            'Amount': f'{amount:.2f}',
            'Type Code': 'OTH'
        })

    # 5 bank fees
    for i in range(5):
        d = base_date + timedelta(days=i * 6)
        balance -= 45.00
        bank_rows.append({
            'Date': fmt_date(d),
            'Description': 'BANK FEE - SERVICE CHARGE',
            'Amount': '-45.00',
            'Running Balance': f'{balance:.2f}'
        })
        dms_rows.append({
            'GL Date': fmt_date(d),
            'Description': 'BANK SERVICE FEE',
            'Reference': f'FEE-{d.strftime("%m%d")}',
            'Amount': '-45.00',
            'Type Code': 'FEE'
        })

    # Shuffle to avoid order bias
    random.shuffle(bank_rows)
    random.shuffle(dms_rows)

    total_bank = len(bank_rows)
    total_dms = len(dms_rows)
    # Expected matchable: 40 checks + 25 deposits + 10 ACH + 5 wires + 3 CVR + 15 misc + 5 fees = 103 pairs
    # Plus ~9 CVR fragments matched via many-to-one
    # Orphans: 15 bank + 10 DMS

    expected = {
        'total_bank': total_bank,
        'total_dms': total_dms,
        'expected_match_rate_min': 80,
        'expected_match_rate_max': 95,
        'note': f'Full month: {total_bank} bank + {total_dms} DMS transactions',
        'expected_orphan_bank': 15,
        'expected_orphan_dms': 10,
    }

    write_scenario('s21', 'Full month simulation', bank_rows, dms_rows, expected)

    print(f"  S21: {total_bank} bank + {total_dms} DMS transactions generated")


def generate_sample_outstanding():
    """Generate a sample outstanding items file for carry-forward testing."""
    filepath = os.path.join(OUTPUT_DIR, 'sample_outstanding.csv')
    with open(filepath, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=[
            'Item ID', 'Source', 'Original Period', 'Transaction Date',
            'Description', 'Amount', 'Check/Reference', 'Type Code',
            'Periods Outstanding', 'Notes'
        ])
        writer.writeheader()
        writer.writerows([
            {'Item ID': 1, 'Source': 'BANK', 'Original Period': '2026-02',
             'Transaction Date': '02/15/2026', 'Description': 'CHECK #4450',
             'Amount': '-890.00', 'Check/Reference': '4450', 'Type Code': 'CHK',
             'Periods Outstanding': 1, 'Notes': 'Outstanding from February'},
            {'Item ID': 2, 'Source': 'DMS', 'Original Period': '2026-02',
             'Transaction Date': '02/20/2026', 'Description': 'PENDING WIRE',
             'Amount': '15000.00', 'Check/Reference': 'WIR-0220', 'Type Code': 'WIR',
             'Periods Outstanding': 1, 'Notes': 'Wire not yet received'},
            {'Item ID': 3, 'Source': 'BANK', 'Original Period': '2026-01',
             'Transaction Date': '01/28/2026', 'Description': 'DEPOSIT',
             'Amount': '2500.00', 'Check/Reference': '', 'Type Code': 'DEP',
             'Periods Outstanding': 2, 'Notes': 'Unidentified deposit'},
        ])


def generate_truist_sample():
    """Generate a Truist-format sample for format detection testing."""
    d = date(2026, 3, 1)
    rows = [
        {'Date': fmt_date(d), 'Description': 'DEPOSIT', 'Debit': '',
         'Credit': '8500.00', 'Balance': '234500.00'},
        {'Date': fmt_date(d), 'Description': 'CHECK 7891', 'Debit': '1100.00',
         'Credit': '', 'Balance': '233400.00'},
        {'Date': fmt_date(d + timedelta(days=1)),
         'Description': 'ACH PAYMENT - INSURANCE', 'Debit': '5670.00',
         'Credit': '', 'Balance': '227730.00'},
        {'Date': fmt_date(d + timedelta(days=2)),
         'Description': 'DEPOSIT', 'Debit': '', 'Credit': '12400.00',
         'Balance': '240130.00'},
    ]
    write_truist_csv('sample_truist.csv', rows)


def main():
    ensure_dir()
    print("Generating ABR test data...")

    generators = [
        generate_s01, generate_s02, generate_s03, generate_s04,
        generate_s05, generate_s06, generate_s07, generate_s08,
        generate_s09, generate_s10, generate_s11, generate_s12,
        generate_s13, generate_s14, generate_s15, generate_s16,
        generate_s17, generate_s18, generate_s19, generate_s20,
        generate_s21,
    ]

    for gen in generators:
        name = gen.__name__.replace('generate_', '').upper()
        print(f"  {name}: {gen.__doc__.strip()}")
        gen()

    generate_sample_outstanding()
    print("  Outstanding items sample generated")

    generate_truist_sample()
    print("  Truist format sample generated")

    print(f"\nAll test data written to: {os.path.abspath(OUTPUT_DIR)}")
    print(f"Total files: {len(os.listdir(OUTPUT_DIR))}")


if __name__ == '__main__':
    main()
