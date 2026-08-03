from __future__ import annotations
from dataclasses import dataclass, field
from io import BytesIO
from typing import Any, Dict, List

import pandas as pd

from planogram.config import SHEET_STORE_LIST, SHEET_STOCK_DISPLAY, SHEET_SO
from planogram.logger import PlanogramLogger


# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------

def _str(val: Any) -> str:
    """Convert a cell value to a stripped string; NaN / None → empty string."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def _clean_store_id(val: Any) -> str:
    """Convert store ID to a clean integer string.

    Handles:
        6349.0  (Excel reads integers as floats)  → "6349"
        "6349 " (trailing whitespace)              → "6349"
        "6349"  (already clean)                    → "6349"
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    try:
        f_val = float(s)
        if f_val.is_integer():
            return str(int(f_val))
    except (ValueError, AttributeError):
        pass
    return s


def _norm(name: str) -> str:
    """Normalise a column header for case-insensitive matching."""
    return str(name).strip().lower()


def _find_col(
    df: pd.DataFrame,
    candidates: List[str],
    fallback_pos: int,
    sheet_name: str,
    logger: PlanogramLogger | None = None,
) -> str:
    """Return the DataFrame column name that matches any of *candidates*.

    Search is case-insensitive.  If no candidate matches, the column at
    *fallback_pos* (0-based) is returned with an info log.

    Notes
    -----
    When an Excel sheet has duplicate column headers (e.g. two columns both
    named "Facings"), pandas automatically renames the second occurrence to
    "Facings.1", the third to "Facings.2", etc.  Pass "Facings.1" as a
    candidate to explicitly target the second "Facings" column.
    """
    cols = list(df.columns)
    norm_map = {_norm(c): c for c in cols}

    for cand in candidates:
        if _norm(cand) in norm_map:
            return norm_map[_norm(cand)]

    # Positional fallback
    if fallback_pos < len(cols):
        actual = cols[fallback_pos]
        if logger:
            logger.info(
                f"[{sheet_name}] Header '{candidates[0]}' not found; "
                f"using column at position {fallback_pos} ('{actual}')."
            )
        return actual

    raise ValueError(
        f"[{sheet_name}] Cannot find column '{candidates[0]}' by name or "
        f"position {fallback_pos}. Available columns: {cols}"
    )


# ---------------------------------------------------------------------------
# WorkbookData container
# ---------------------------------------------------------------------------

@dataclass
class WorkbookData:
    """Holds the three parsed DataFrames and their resolved column maps."""

    store_list:     pd.DataFrame
    stock_display:  pd.DataFrame
    special_orders: pd.DataFrame
    cols_sl: Dict[str, str] = field(default_factory=dict)
    cols_sd: Dict[str, str] = field(default_factory=dict)
    cols_so: Dict[str, str] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Load workbook from bytes
# ---------------------------------------------------------------------------

def load_workbook_from_bytes(
    file_bytes: bytes,
    logger: PlanogramLogger | None = None,
) -> WorkbookData:
    """Load all three input sheets from an in-memory Excel workbook.

    Parameters
    ----------
    file_bytes:
        Raw bytes of the .xlsx file (from Streamlit's file_uploader).
    logger:
        PlanogramLogger instance for validation messages.

    Returns
    -------
    WorkbookData
        Parsed DataFrames with resolved column name maps.
    """
    try:
        xl = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    except Exception as exc:
        raise RuntimeError(f"Cannot open workbook: {exc}") from exc

    def _get_sheet(desired: str) -> str:
        """Case-insensitive sheet name lookup."""
        for name in xl.sheet_names:
            if name.strip().lower() == desired.strip().lower():
                return name
        raise ValueError(
            f"Sheet '{desired}' not found in workbook. "
            f"Available sheets: {xl.sheet_names}"
        )

    s_sl = _get_sheet(SHEET_STORE_LIST)
    s_sd = _get_sheet(SHEET_STOCK_DISPLAY)
    s_so = _get_sheet(SHEET_SO)

    # Read everything as strings to avoid Excel numeric/date type issues.
    df_sl = xl.parse(s_sl, dtype=str)
    df_sd = xl.parse(s_sd, dtype=str)
    df_so = xl.parse(s_so, dtype=str)

    for df in (df_sl, df_sd, df_so):
        # Strip whitespace from headers.
        # Pandas auto-disambiguates duplicate headers: "Facings" → "Facings.1"
        df.columns = [str(c).strip() for c in df.columns]
        df.dropna(how="all", inplace=True)
        df.reset_index(drop=True, inplace=True)

    # -----------------------------------------------------------------------
    # Sheet 1 — Store List
    # Expected column order: Store | Current Store POG | Current LFT | Notes
    # -----------------------------------------------------------------------
    cols_sl = {
        "store": _find_col(df_sl, ["Store"], 0, s_sl, logger),
        "pog":   _find_col(
                     df_sl,
                     ["Current Store POG", "Store POG", "Current Pog", "POG"],
                     1, s_sl, logger,
                 ),
        "lft":   _find_col(
                     df_sl,
                     ["Current LFT", "LFT", "Current Lft"],
                     2, s_sl, logger,
                 ),
        "notes": _find_col(df_sl, ["Notes", "Note"], 3, s_sl, logger),
    }

    # -----------------------------------------------------------------------
    # Sheet 2 — Stock SKUs and Displays
    #
    # Typical layout (columns A–J):
    #   A (0) Store
    #   B (1) Stock SKU
    #   C (2) Stock Description
    #   D (3) [blank / sort col]
    #   E (4) Facings  ← Stock Facing
    #   F (5) Display SKU
    #   G (6) Display Description
    #   H (7) [blank / sort col]
    #   I (8) Facings  ← Display Facing  (pandas renames to "Facings.1")
    #   J (9) CF
    #
    # The duplicate-header rename is KEY: if both stock and display facing
    # columns are labelled "Facings", pandas produces "Facings" (stock) and
    # "Facings.1" (display).  We exploit this for reliable detection.
    # -----------------------------------------------------------------------
    sd_cols = list(df_sd.columns)
    cols_sd = {
        "store":      _find_col(df_sd, ["Store"], 0, s_sd, logger),
        "stock_sku":  _find_col(
                          df_sd,
                          ["Stock SKU", "Stock Sku", "SKU"],
                          1, s_sd, logger,
                      ),
        "stock_desc": _find_col(
                          df_sd,
                          ["Stock Description", "Stock Desc", "Description"],
                          2, s_sd, logger,
                      ),
        # First "Facings" column (position E / index 4)
        "stock_face": _find_col(
                          df_sd,
                          ["Facings", "Facing", "Stock Facings", "Stock Facing"],
                          4, s_sd, logger,
                      ),
        "disp_sku":   _find_col(
                          df_sd,
                          ["Display SKU", "Display Sku"],
                          5, s_sd, logger,
                      ),
        "disp_desc":  _find_col(
                          df_sd,
                          ["Display Description", "Display Desc"],
                          6, s_sd, logger,
                      ),
        # Second "Facings" column — pandas renames it to "Facings.1" (position I / index 8)
        "disp_face":  _find_col(
                          df_sd,
                          ["Facings.1", "Facing.1", "Display Facings", "Display Facing"],
                          8, s_sd, logger,
                      ),
        "cf":         _find_col(df_sd, ["CF"], len(sd_cols) - 1, s_sd, logger),
    }

    # -----------------------------------------------------------------------
    # Sheet 3 — Special Order Boards
    #
    # Typical layout:
    #   A (0) Store
    #   B (1) SO SKU (may be labelled "Display SKU" or "SKU")
    #   C (2) Description
    #   D (3) [blank / sort col]
    #   E (4) Facings
    #   ... (additional cols)
    #   last  CF
    # -----------------------------------------------------------------------
    so_cols = list(df_so.columns)
    cols_so = {
        "store":   _find_col(df_so, ["Store"], 0, s_so, logger),
        "so_sku":  _find_col(
                       df_so,
                       ["SO SKU", "So Sku", "Display SKU", "SKU"],
                       1, s_so, logger,
                   ),
        "so_desc": _find_col(
                       df_so,
                       ["Description", "Display Description", "SO Description"],
                       2, s_so, logger,
                   ),
        "so_face": _find_col(
                       df_so,
                       ["Facings", "Facing"],
                       4, s_so, logger,
                   ),
        "cf":      _find_col(df_so, ["CF"], len(so_cols) - 1, s_so, logger),
    }

    return WorkbookData(
        store_list=df_sl,
        stock_display=df_sd,
        special_orders=df_so,
        cols_sl=cols_sl,
        cols_sd=cols_sd,
        cols_so=cols_so,
    )
