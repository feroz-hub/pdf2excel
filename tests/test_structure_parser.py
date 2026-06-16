"""Tests for the structure parser state machine."""

from __future__ import annotations

import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from standard_profiles import Nist80053Profile
from structure_parser import parse_blocks_to_items
import validation


def _block(text: str, page: int = 1, block_type: str = "paragraph") -> dict:
    return {
        "page": page,
        "block_type": block_type,
        "text": text,
        "bbox": [50, 100, 500, 120],
        "confidence": 0.95,
        "issues": [],
    }


def _exported_items(blocks):
    items, _rej = parse_blocks_to_items(
        blocks, profile=Nist80053Profile, include_front_matter=False
    )
    items = validation.validate_items(items, profile_name="nist80053")
    return validation.filter_exportable_items(items, profile_name="nist80053")


class TestFrontMatterRejected:
    def test_example_a_not_exported(self):
        blocks = [
            _block("JOINT TASK FORCE"),
            _block("Authority"),
            _block("Abstract"),
        ]
        exported = _exported_items(blocks)
        assert exported == []


class TestAC1Grouping:
    def test_example_b_single_item(self):
        blocks = [
            _block("AC-1 POLICY AND PROCEDURES"),
            _block("Control:\na. Develop, document..."),
            _block("Discussion: Policy discussion here."),
            _block("Related Controls: AC-2, AC-3."),
            _block("References: NIST documents."),
        ]
        exported = _exported_items(blocks)
        assert len(exported) == 1
        it = exported[0]
        assert it["clause_id"] == "AC-1"
        assert it["title"] == "POLICY AND PROCEDURES"
        assert "Control:" in it["text"]
        assert "Discussion:" in it["text"]
        assert "Related Controls:" in it["text"]
        assert "References:" in it["text"]


class TestAC2Enhancements:
    def test_example_c_three_items(self):
        blocks = [
            _block("AC-2 ACCOUNT MANAGEMENT"),
            _block("Control:\nManage accounts."),
            _block("(1) ACCOUNT MANAGEMENT | AUTOMATED SYSTEM ACCOUNT MANAGEMENT"),
            _block("Support the management of system accounts using [Assignment: organization-defined automated mechanisms]."),
            _block("Discussion: Automated system account management includes..."),
            _block("Related Controls: None."),
            _block("(2) ACCOUNT MANAGEMENT | AUTOMATED TEMPORARY AND EMERGENCY ACCOUNT MANAGEMENT"),
            _block("Automatically remove temporary accounts..."),
        ]
        exported = _exported_items(blocks)
        ids = [it["clause_id"] for it in exported]
        assert ids == ["AC-2", "AC-2(1)", "AC-2(2)"]
        ac2_1 = next(it for it in exported if it["clause_id"] == "AC-2(1)")
        assert "Discussion:" in ac2_1["text"]
        assert "Related Controls:" in ac2_1["text"]
        assert not any(it["clause_id"] == "" for it in exported)


class TestNoiseRejected:
    def test_example_d_not_exported(self):
        blocks = [
            _block("NIST SP 800-53, R EV . 5 S ECURITY AND P RIVACY C ONTROLS"),
            _block("This publication is available free of charge from:"),
            _block("https://doi.org/10.6028/NIST.SP.800-53r5"),
            _block("CHAPTER THREE PAGE 20"),
        ]
        exported = _exported_items(blocks)
        assert exported == []


class TestMixedCaseNumberedSubClauses:
    """Numbered sub-clauses in control text must not become enhancements."""

    def test_lowercase_numbered_items_stay_in_control(self):
        blocks = [
            _block("AC-3 ACCESS ENFORCEMENT"),
            _block("Control:"),
            _block("(1) Changing one or more security attributes on subjects,"),
            _block("objects, the system, or system components;"),
            _block("(2) Choosing the security attributes for new objects;"),
            _block("(1) RESTRICTED ACCESS TO PRIVILEGED FUNCTIONS"),
            _block("[Withdrawn: Incorporated into AC-6.]"),
        ]
        exported = _exported_items(blocks)
        ids = [it["clause_id"] for it in exported]
        assert ids.count("AC-3") == 1
        assert "AC-3(1)" in ids
        base = next(it for it in exported if it["clause_id"] == "AC-3")
        assert "Changing one or more security attributes" in base["text"]
        assert "Choosing the security attributes" in base["text"]
