# ABR Tool Rebuild — Project Context for Claude Code

## What This Project Is

Auto Bank Reconciliation (ABR) tool rebuild for Jim Coleman Automotive — a 7-location
dealership group in the Maryland/DC area. The original tool was a custom Excel/VBA
application used by controllers at all 7 stores to reconcile bank statements against
the Reynolds & Reynolds (R&R) DMS every month. The original developer passed away in
2024 and all licenses expired. Dealerships are now doing manual reconciliation, which
is slower and error-prone. This project rebuilds it — better, not just cloned.

**Owner:** Chris Whitney (controller / operator)
**Platform:** Excel/VBA (non-negotiable — this is what controllers use daily)
**Banks:** Bank of America + Truist
**DMS:** Reynolds & Reynolds (R&R)
**Locations:** 7 Jim Coleman stores

---

## Spec-First Discipline — Non-Negotiable

**Zero code is written until the functional spec is complete and signed off.**

This is not a suggestion. Every time we've skipped spec work and gone straight to
building, we've had to tear it out. The spec defines what the tool must do, must NOT
do, and how it handles every edge case. Only after that do we design algorithms.
Only after algorithm design do we write VBA.

Current phase: **Phase 1 — Writing the Functional Spec**

---

## Build Sequence (Do Not Skip Steps)

| Phase | Name | Status |
|-------|------|--------|
| 0 | Intelligence Gathering | ✅ COMPLETE |
| 1 | Write Functional Spec | 🔴 CURRENT BLOCKER |
| 2 | Design Matching Algorithm | Locked |
| 3 | Design User Confirmation Workflow | Locked |
| 4 | Design Split / Many-to-One Logic (CVR) | Locked |
| 5 | Build Excel/VBA Prototype (1 location) | Locked |
| 6 | Test Against Historical Data | Locked |
| 7 | Roll Out to All 7 Locations | Locked |

---

## What We Know About the Original Tool (Fully Reverse-Engineered)

Phase 0 is complete. We analyzed all available historical reconciliation files, bank
statements, R&R DMS GL exports, outstanding items files, and a 23-minute Zoom screen
recording of the live tool being used by a controller in 2023. Full behavior has been
reconstructed.

### Original Matching Algorithm

- **Primary match:** Exact dollar amount
- **Secondary validation:** Check number extraction + date proximity
- **Auto-match rate:** ~82–84% on first import of each month
- **Workflow cadence:** Auto-reconciliation ONLY runs on the first import of each month.
  All subsequent sessions in that month use a manual "Update Recon" mode where the
  controller matches transactions one at a time.

### Original UI (5-Step Main Menu)

The original had a 5-step main menu workflow:
1. Import bank statement
2. Import DMS data
3. Run auto-reconciliation
4. Manual match / review unmatched items
5. Finalize / export

Manual matching screen showed real-time completion percentages as items were cleared.

### Data Sources

**Bank statements** — imported from Bank of America and Truist export files.
Contains: transaction date, description, amount, check number (where applicable).

**DMS GL exports** — exported from Reynolds & Reynolds. Contains: GL date, description,
reference number, amount, transaction type code.

**Outstanding items files** — carried-forward unmatched items from prior periods.

---

## The Four Critical Flaws (Must Fix in Rebuild)

These are not nice-to-haves. These are the reason we're rebuilding instead of cloning.

### Flaw 1: False Matches from Exact-Dollar-Only Logic
The original used exact dollar amount as its primary match criterion with insufficient
secondary verification. Dollar amounts are not unique — two different transactions can
have the same dollar amount, especially for round numbers, recurring fees, or payroll
runs. The original would match the wrong pair and commit it without flagging. This
produced silent errors that only surfaced during month-end review.

**Fix:** Weighted multi-factor confidence scoring. Amount match alone is never
sufficient to auto-commit. Confidence must incorporate check number, date proximity,
and transaction type. Low-confidence matches surface for human review.

### Flaw 2: CVR Many-to-One Matching Failure
CVR (Customer Vehicle Receivable) transactions are the single most painful recurring
issue for controllers. In R&R DMS, a CVR often appears as a single lump-sum entry.
But on the bank side, that same CVR hits as multiple separate deposits — the bank
fragments what the DMS shows as one. The original had no logic to handle this. It
either failed to match at all or matched one fragment to the full DMS amount (wrong).

This is the #1 pain point reported by users. Must be solved, not worked around.

**Fix:** Dedicated many-to-one matching mode. Group multiple bank-side transactions
and match them against a single DMS entry. The sum of the group must equal the DMS
amount within a defined tolerance. Special UI for reviewing and confirming these grouped
matches.

### Flaw 3: Auto-Clearing Without User Confirmation
The original auto-committed matches during the auto-reconciliation pass without
showing them to the user first. Once committed, they were treated as reconciled. If
a false match was committed, you had to manually find and un-clear it — which required
knowing it was wrong in the first place.

**Fix:** No match is ever auto-committed. Every match — even high-confidence ones —
is staged for review. The controller sees what the engine proposes, approves or
rejects each one, and only then does it commit. Auto-reconciliation becomes
auto-SUGGESTION, not auto-COMMITMENT.

### Flaw 4: No Split-Transaction Handling
Related to Flaw 2 but broader. The original couldn't handle any fragmented
transactions in either direction — bank splitting what DMS shows as one, or DMS
splitting what the bank shows as one. These just ended up in the unmatched pile
every month for manual handling.

**Fix:** First-class split-transaction support in both directions, with proper
UI for grouping and confirming splits.

---

## Confidence Scoring Design Principles

The new matching algorithm must produce a confidence score (0–100%) on every
proposed match, not a binary pass/fail. Rules:

- **Exact amount + exact check number + date within 1 day** → very high confidence
- **Exact amount + date within 3 days** → moderate confidence
- **Exact amount only, no check number, date window > 5 days** → low confidence
- **Any low-confidence match MUST surface for human review — never auto-stage it**
- **Confidence thresholds are configurable** — controllers at different locations may
  have different risk tolerances. The spec must define defaults and the mechanism for
  adjusting them.

---

## Audit Trail Requirements

The original had zero audit trail. The rebuild must log:

- Who initiated each reconciliation session
- Every match that was proposed by the engine
- Every match that was accepted by the controller (with timestamp + user)
- Every match that was rejected by the controller (with reason if provided)
- Every manual match made outside the auto-suggestion workflow
- Every split/group configuration made by the controller

This is non-negotiable. Month-end review, external audits, and inter-location
troubleshooting all depend on knowing what happened and who did it.

---

## Architecture Decisions (To Be Finalized in Spec)

### Open Question: Single Workbook vs. Per-Location Instances
- **Option A:** One shared workbook, location selected at runtime via dropdown
- **Option B:** 7 separate identical workbook instances, one per location
- **Tradeoffs:** Option A = easier updates, harder permissions. Option B = simpler
  per-controller ownership, update distribution overhead.
- **Decision:** TBD in spec. Do not assume either during prototype phase.

### Prototype Target: Single Location First
Phase 5 builds for one location only. Do not attempt to build multi-location
architecture until the single-location prototype is validated against historical data.

---

## Success Criteria

The rebuild is successful when:
1. Auto-suggestion rate **meets or exceeds** the original 82–84% baseline
2. Zero auto-committed false matches (the engine never commits anything without
   controller confirmation)
3. CVR many-to-one scenarios are handled — no more manual pile for these
4. Complete audit trail exists for every reconciliation session
5. All 7 locations can use it without requiring developer support

---

## Technical Environment

- **Platform:** Excel/VBA — no Python, no web app, no database. Controllers open
  Excel. Period. This constraint is absolute.
- **DMS:** Reynolds & Reynolds (R&R) — exports GL data in fixed formats
- **Banks:** Bank of America + Truist — each has its own export format. The tool
  must handle both, and the formats are different.
- **Deployment:** Each controller's machine. No server, no shared network execution.
- **Future state (not in scope for initial build):** Power BI reporting layer for
  cross-location visibility. Design for it, but don't build it yet.

---

## What NOT to Do

- **Do not write any code before the functional spec is signed off.** If you find
  yourself writing VBA before the spec document exists, stop.
- **Do not clone the original.** We are fixing four known flaws. The original's
  behavior is reference material, not a blueprint.
- **Do not auto-commit any match.** Ever. Not even 99% confidence. Stage it.
- **Do not assume CVR fragmentation is an edge case.** It is the most common
  recurring pain point. Treat it as a first-class scenario.
- **Do not skip the audit trail.** It is not optional scope.
- **Do not build multi-location before single-location is validated.**
- **Do not use third-party VBA libraries or add-ins** that require separate
  installation. The tool must work on any Windows machine with Excel. No dependencies.

---

## Context Document Location

Full project history and session logs are maintained in Google Drive:
- **Active Projects Doc ID:** `1J1WiHa3Q1SwNCfaI9geKFaxRp-gnZGgZlo2NVmjVHLA`
- **Master Context Doc ID:** `1U9MgxXX5kvLrrzh7CINNHL4wd-fUlS5dlQ0pi90zwhY`

These contain complete session history, all decisions made to date, and the full
reverse-engineering findings from Phase 0. If you need more context than what's in
this file, those are the authoritative sources.

---

## Current Session Starting Point

If you're reading this at the start of a new Claude Code session:

1. Phase 0 (intelligence gathering) is complete
2. We are in Phase 1 — writing the functional spec
3. No code exists yet
4. The next concrete task is drafting the functional spec document covering:
   - Input data formats (BofA, Truist, R&R DMS)
   - Matching rules and confidence scoring logic
   - User confirmation workflow
   - Split/many-to-one transaction handling
   - Audit trail schema
   - Error and exception handling
   - Location architecture decision

Ask Chris which part of the spec to start on before beginning any drafting.
