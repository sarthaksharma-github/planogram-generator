"""
Comprehensive test suite for HD Planogram Generator and ChaiUI compliance.
"""
import io
import re
import pytest
import pandas as pd
from openpyxl import Workbook, load_workbook

from app import (
    BayRule,
    BAY_RULES,
    SKURecord,
    parse_pog_string,
    _parse_facing,
    apply_notes_rules,
    audit_planogram_df,
    write_output_bytes,
    generate_store_sku_sheet,
    PlanogramLogger,
    allocate_bay,
    allocate_store,
)


def test_pog_parser_variants():
    """Verify all POG parser format variants."""
    # FORMAT C: single repeated size
    assert parse_pog_string("12 Bay - 99") == ["99"] * 12
    assert parse_pog_string("12 Bay - 99 - Floor Tile Test - Str 1602 - June 26") == ["99"] * 12
    assert parse_pog_string("1 Bay 99 - OPP Plug & Play") == ["99"]

    # FORMAT B: explicit comma list
    assert parse_pog_string("7 BAY - 99,99,99,99,99,99,87") == ["99"] * 6 + ["87"]
    assert parse_pog_string("10 BAY - 99C,99,99,99,99,99C,99,99,99,99") == [
        "99C", "99", "99", "99", "99", "99C", "99", "99", "99", "99"
    ]

    # FORMAT A: multiplier shorthand
    assert parse_pog_string("10 Bay - 8-99,2-51") == ["99"] * 8 + ["51"] * 2
    assert parse_pog_string("5 Bay- 4-99, 1-87 - Stone Look Tile") == ["99"] * 4 + ["87"]

    # Invalid / empty inputs
    assert parse_pog_string("") == []
    assert parse_pog_string("Invalid POG") == []


def test_facing_clamping():
    """Verify facing values are clamped to 1 or 2."""
    logger = PlanogramLogger()
    assert _parse_facing(1, "1001", "SKU1", logger) == 1
    assert _parse_facing("2", "1001", "SKU1", logger) == 2
    assert _parse_facing(3, "1001", "SKU1", logger) == 1  # clamped to 1
    assert _parse_facing("invalid", "1001", "SKU1", logger) == 1  # invalid clamped to 1
    assert _parse_facing(0, "1001", "SKU1", logger) == 1  # clamped to 1


def test_notes_rules_partition():
    """Verify notes rules partition matching SKUs stably to the end."""
    records = [
        SKURecord(sku="101", description="Regular Tile Grey", facing=1, sku_type="Stock"),
        SKURecord(sku="102", description="Vigo Marble White", facing=1, sku_type="Stock"),
        SKURecord(sku="103", description="Baja Sand Brown", facing=1, sku_type="Stock"),
        SKURecord(sku="104", description="Standard Granite", facing=1, sku_type="Stock"),
    ]

    # Without matching note
    res_no_rule = apply_notes_rules(records, "Standard store flow")
    assert [r.sku for r in res_no_rule] == ["101", "102", "103", "104"]

    # With matching note "Force Baja & Vigos to the end"
    res_ruled = apply_notes_rules(records, "Force Baja & Vigos to the end")
    assert [r.sku for r in res_ruled] == ["101", "104", "102", "103"]


def test_audit_planogram_df():
    """Verify audit_planogram_df logic on compliant vs non-compliant stores."""
    stores_df = pd.DataFrame({
        "Store": ["1001", "1002", "1003"],
        "Notes": [
            "Force Baja & Vigos to the end",
            "Normal layout",
            "Move Baja to end",
        ]
    })

    # Store 1001: Pass (Vigo is at the end)
    # Store 1003: Fail (Baja is in front of regular SKU)
    pog_df = pd.DataFrame({
        "Store": ["1001", "1001", "1003", "1003"],
        "SKU Type": ["Stock", "Stock", "Stock", "Stock"],
        "SKU Description": [
            "Regular Grey Tile",
            "Vigo White Marble",
            "Baja Sand Brown",
            "Regular White Tile",
        ],
    })

    audit_result = audit_planogram_df(stores_df, pog_df, notes_filter="baja|vigo", keywords=["BAJA", "VIGO"], categories=["Stock"])

    assert len(audit_result) == 2  # Only 1001 and 1003 have matching notes
    store_1001 = audit_result[audit_result["Store"] == "1001"].iloc[0]
    store_1003 = audit_result[audit_result["Store"] == "1003"].iloc[0]

    assert bool(store_1001["Passed"]) is True
    assert store_1001["Status"] == "PASS"

    assert bool(store_1003["Passed"]) is False
    assert store_1003["Status"] == "FAIL"


def test_write_output_bytes():
    """Verify openpyxl writer bundles Planogram, Validation, and Audit Report sheets."""
    # Create a minimal base workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Store List"
    ws.append(["Store", "Current Store POG", "Current LFT", "Notes"])
    ws.append(["1001", "1 Bay - 99", "None", "Normal"])

    bio = io.BytesIO()
    wb.save(bio)
    raw_bytes = bio.getvalue()

    pog_df = pd.DataFrame([{
        "Store": "1001", "Bay#": 1, "Bay Size": "99",
        "Shelf": 1, "Position": 1, "SKU": "999",
        "SKU Type": "Display", "SKU Description": "Sample", "Facing": 1
    }])
    val_df = pd.DataFrame([{"Level": "INFO", "Store": "1001", "Bay#": 1, "Message": "Test"}])
    audit_df = pd.DataFrame([{"Store": "1001", "Status": "PASS", "Details": "All pass"}])

    out_bytes = write_output_bytes(raw_bytes, pog_df, val_df, audit_df)
    assert len(out_bytes) > 0

    out_wb = load_workbook(io.BytesIO(out_bytes))
    assert "Generated Planogram" in out_wb.sheetnames
    assert "Validation" in out_wb.sheetnames
    assert "Audit Report" in out_wb.sheetnames


def test_store_sku_generation():
    """Verify store SKU sheet generator."""
    wb = Workbook()
    ws_store = wb.active
    ws_store.title = "Store"
    ws_store.append(["Store", "SO Board Set"])
    ws_store.append(["1001", "1 Bay South"])
    ws_store.append(["1002", "1 Bay North"])

    ws_sku = wb.create_sheet(title="SKU")
    ws_sku.append(["SKU", "Description", "1 Bay South", "1 Bay North"])
    ws_sku.append(["SKU-A", "Tile A", "Y", ""])
    ws_sku.append(["SKU-B", "Tile B", "", "Y"])

    bio = io.BytesIO()
    wb.save(bio)
    raw_bytes = bio.getvalue()

    result_bytes = generate_store_sku_sheet(raw_bytes)
    res_wb = load_workbook(io.BytesIO(result_bytes))
    assert "Store SKU" in res_wb.sheetnames

    ws_out = res_wb["Store SKU"]
    rows = list(ws_out.iter_rows(values_only=True))
    assert rows[0] == ("Store", "SKU", "Description", "Store-SKU", "Facings", "SKU - Facing")
    assert len(rows) == 3  # Header + 2 data rows


def test_chaiui_compliance_rules():
    """Automated ChaiUI design review audit checks on app.py."""
    with open("app.py", "r", encoding="utf-8") as f:
        code = f.read()

    # Rule 1: Zero em dashes anywhere (code, comments, copy)
    assert "—" not in code, "Em dash character found in app.py"
    assert "\u2014" not in code, "Unicode em dash found in app.py"
    assert "&mdash;" not in code, "HTML em dash entity found in app.py"

    # Rule 2: Zero orange fill colors anywhere in styling
    orange_fill_patterns = [
        r"bg-orange-",
        r"bg-\[#(FF7D0C|ff7d0c|FE9332|fe9332)\]",
        r"#F96302",
        r"#f96302",
        r"#cc4f00",
        r"#CC4F00",
        r"linear-gradient\(.*#F96302",
        r"linear-gradient\(.*#cc4f00",
    ]
    for pattern in orange_fill_patterns:
        match = re.search(pattern, code, re.IGNORECASE)
        assert match is None, f"Prohibited orange fill pattern matched: {pattern}"

    # Rule 3: Zero box-shadows or drop-shadows (except box-shadow: none !important)
    shadow_lines = [
        line.strip() for line in code.splitlines()
        if "box-shadow" in line or "drop-shadow" in line or "text-shadow" in line
    ]
    for line in shadow_lines:
        assert "none" in line, f"Shadow effect found that is not none: {line}"

    # Rule 4: Zero uppercase eyebrows
    assert "tracking-widest" not in code
    assert "text-transform: uppercase" not in code

    # Rule 5: Dark background #0a0a0a present
    assert "#0a0a0a" in code

    # Rule 6: Cream highlight (#ffedd5) present
    assert "#ffedd5" in code

    # Rule 7: Manrope and Montserrat fonts referenced
    assert "Manrope" in code
    assert "Montserrat" in code

    # Rule 8: Loading screen present with 'HD Planogram Generator'
    assert "splash-screen" in code
    assert "HD Planogram Generator" in code
