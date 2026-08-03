from __future__ import annotations
from typing import Any, Dict, List

from planogram.models import BayRule

# ---------------------------------------------------------------------------
# BAY_RULES
# To add a new bay type: add one row here. No other code needs to change.
# ---------------------------------------------------------------------------
BAY_RULES: Dict[str, BayRule] = {
    "99":  BayRule(display=6, so=2, stock=6),
    "99C": BayRule(display=4, so=2, stock=4),
    "87":  BayRule(display=4, so=1, stock=4),
    "75":  BayRule(display=4, so=1, stock=4),
    "51":  BayRule(display=2, so=1, stock=2),
}

# ---------------------------------------------------------------------------
# NOTES_RULES
# To add a new Notes rule: append one dict here. No other code needs to change.
# ---------------------------------------------------------------------------
NOTES_RULES: List[Dict[str, Any]] = [
    {
        "trigger": "force baja & vigos to the end",
        "keywords": ["baja", "vigo"],
    },
]

# ---------------------------------------------------------------------------
# Sheet names (matched case-insensitively at load time)
# ---------------------------------------------------------------------------
SHEET_STORE_LIST    = "Store List"
SHEET_STOCK_DISPLAY = "Stock SKUs and Displays"
SHEET_SO            = "Special Order Boards"

# ---------------------------------------------------------------------------
# Output sheet names
# ---------------------------------------------------------------------------
OUTPUT_SHEET_PLANOGRAM  = "Generated Planogram"
OUTPUT_SHEET_VALIDATION = "Validation"

# ---------------------------------------------------------------------------
# Output columns for the Generated Planogram sheet
# ---------------------------------------------------------------------------
PLANOGRAM_COLUMNS: List[str] = [
    "Store", "Bay#", "Bay Size", "Shelf", "Position",
    "SKU", "SKU Type", "SKU Description", "Facing",
]

# ---------------------------------------------------------------------------
# Shelf numbering
# ---------------------------------------------------------------------------
SHELF_DISPLAY      = 1  # Shelf 1  -> Display Boards
SHELF_SO           = 2  # Shelf 2  -> Special Order Boards
SHELF_STOCK_FIRST  = 3  # Shelf 3  -> First half of Stock
SHELF_STOCK_SECOND = 4  # Shelf 4  -> Remaining Stock

# ---------------------------------------------------------------------------
# LFT sentinel values — these mean "no LFT for this store"
# ---------------------------------------------------------------------------
LFT_IGNORE_VALUES: frozenset = frozenset(["-", "", "none", "null", "n/a", "na"])

# ---------------------------------------------------------------------------
# Valid facing values — anything else is clamped to 1.
# This prevents expand_facing() exploding when a CF or other numeric column
# is accidentally mapped to the Facing column.
# ---------------------------------------------------------------------------
VALID_FACINGS: frozenset = frozenset([1, 2])
