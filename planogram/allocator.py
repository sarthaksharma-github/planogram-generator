from __future__ import annotations
import math
from collections import defaultdict
from typing import Any, Dict, List, Tuple

import pandas as pd

from planogram.config import (
    BAY_RULES,
    NOTES_RULES,
    PLANOGRAM_COLUMNS,
    LFT_IGNORE_VALUES,
    VALID_FACINGS,
    SHELF_DISPLAY,
    SHELF_SO,
    SHELF_STOCK_FIRST,
    SHELF_STOCK_SECOND,
)
from planogram.loader import WorkbookData, _clean_store_id, _str
from planogram.logger import PlanogramLogger
from planogram.models import BayRule, SKURecord
from planogram.parser import parse_pog_string


# ---------------------------------------------------------------------------
# Facing parser — clamps invalid values to 1
# ---------------------------------------------------------------------------

def _parse_facing(
    raw: Any,
    store: Any,
    sku: Any,
    logger: PlanogramLogger,
) -> int:
    """Convert a raw Facing cell value to int (must be 1 or 2).

    CRITICAL: Always returns 1 or 2.  Never returns the raw cell value when it
    is outside {1, 2}.  This prevents expand_facing() from generating dozens
    or hundreds of copies of a SKU when a CF or other large numeric column is
    accidentally mapped to the Facing column.
    """
    try:
        val = int(float(str(raw).strip()))
    except (ValueError, TypeError):
        logger.warning(
            f"Invalid Facing '{raw}' for SKU {sku} — defaulting to 1.",
            store=store,
        )
        return 1

    if val not in VALID_FACINGS:
        logger.warning(
            f"Facing value '{raw}' for SKU {sku} is not in {{1, 2}} — defaulting to 1.",
            store=store,
        )
        return 1  # Must return 1, NOT val — prevents expansion explosion

    return val


# ---------------------------------------------------------------------------
# SKU index builders
# ---------------------------------------------------------------------------

def build_display_index(
    df: pd.DataFrame,
    cols: Dict[str, str],
    logger: PlanogramLogger,
) -> Dict[str, List[SKURecord]]:
    """Build store → [SKURecord] mapping for Display Board SKUs."""
    index: Dict[str, List[SKURecord]] = defaultdict(list)
    for _, row in df.iterrows():
        store  = _clean_store_id(row[cols["store"]])
        sku    = _str(row[cols["disp_sku"]])
        desc   = _str(row[cols["disp_desc"]])
        facing = _parse_facing(row[cols["disp_face"]], store, sku, logger)
        if not store or not sku:
            continue
        index[store].append(
            SKURecord(sku=sku, description=desc, facing=facing, sku_type="Display")
        )
    return dict(index)


def build_stock_index(
    df: pd.DataFrame,
    cols: Dict[str, str],
    logger: PlanogramLogger,
) -> Dict[str, List[SKURecord]]:
    """Build store → [SKURecord] mapping for Stock SKUs."""
    index: Dict[str, List[SKURecord]] = defaultdict(list)
    for _, row in df.iterrows():
        store  = _clean_store_id(row[cols["store"]])
        sku    = _str(row[cols["stock_sku"]])
        desc   = _str(row[cols["stock_desc"]])
        facing = _parse_facing(row[cols["stock_face"]], store, sku, logger)
        if not store or not sku:
            continue
        index[store].append(
            SKURecord(sku=sku, description=desc, facing=facing, sku_type="Stock")
        )
    return dict(index)


def build_so_index(
    df: pd.DataFrame,
    cols: Dict[str, str],
    logger: PlanogramLogger,
) -> Dict[str, List[SKURecord]]:
    """Build store → [SKURecord] mapping for Special Order Boards."""
    index: Dict[str, List[SKURecord]] = defaultdict(list)
    for _, row in df.iterrows():
        store  = _clean_store_id(row[cols["store"]])
        sku    = _str(row[cols["so_sku"]])
        desc   = _str(row[cols["so_desc"]])
        facing = _parse_facing(row[cols["so_face"]], store, sku, logger)
        if not store or not sku:
            continue
        index[store].append(
            SKURecord(sku=sku, description=desc, facing=facing, sku_type="SO")
        )
    return dict(index)


# ---------------------------------------------------------------------------
# Facing expansion
# ---------------------------------------------------------------------------

def expand_facing(records: List[SKURecord]) -> List[SKURecord]:
    """Expand Facing > 1 so each returned record represents one position.

    A SKU with Facing=2 produces two consecutive identical records.
    Because _parse_facing() always returns 1 or 2, this is safe and bounded.
    """
    expanded: List[SKURecord] = []
    for rec in records:
        for _ in range(max(1, rec.facing)):
            expanded.append(rec)
    return expanded


# ---------------------------------------------------------------------------
# Notes rules (stable partition)
# ---------------------------------------------------------------------------

def apply_notes_rules(
    records: List[SKURecord],
    notes: str,
) -> List[SKURecord]:
    """Apply matching NOTES_RULES via stable partition.

    For each triggered rule:
      - Checks if any trigger keyword appears in the Notes text (case-insensitive).
      - SKUs whose description contains any SKU keyword (e.g. "baja", "vigo") -> moved to tail.
      - All others -> remain in head.
      - Relative order is preserved within both head and tail.
    """
    if not notes or not isinstance(notes, str):
        return records

    notes_lower = notes.strip().lower()
    result = list(records)

    for rule in NOTES_RULES:
        # Support both new trigger_keywords and legacy trigger string
        triggers = rule.get("trigger_keywords", [])
        if "trigger" in rule:
            triggers.append(rule["trigger"])

        # Check if ANY trigger keyword appears in the notes text
        is_triggered = any(tr.lower() in notes_lower for tr in triggers if tr)
        if not is_triggered:
            continue

        # Support both new sku_keywords and legacy keywords list
        sku_kws = [kw.lower() for kw in rule.get("sku_keywords", rule.get("keywords", [])) if kw]
        if not sku_kws:
            continue

        head: List[SKURecord] = []
        tail: List[SKURecord] = []
        for rec in result:
            rec_desc_lower = rec.description.lower()
            if any(kw in rec_desc_lower for kw in sku_kws):
                tail.append(rec)
            else:
                head.append(rec)
        result = head + tail

    return result


# ---------------------------------------------------------------------------
# Bay allocator helpers
# ---------------------------------------------------------------------------

def _consume(
    pool: List[SKURecord],
    pointer: List[int],
    count: int,
    store: Any,
    bay_num: int,
    sku_type: str,
    logger: PlanogramLogger,
) -> List[SKURecord]:
    """Take *count* records from *pool* starting at *pointer[0]*."""
    start = pointer[0]
    end   = start + count
    taken = pool[start:end]
    if len(taken) < count:
        logger.warning(
            f"Not enough {sku_type} SKUs for Bay {bay_num}: "
            f"need {count}, have {len(taken)} remaining.",
            store=store,
            bay_num=bay_num,
        )
    pointer[0] = min(end, len(pool))
    return taken


def allocate_bay(
    store: Any,
    bay_num: int,
    bay_size: str,
    rule: BayRule,
    disp_pool: List[SKURecord],
    so_pool: List[SKURecord],
    stock_pool: List[SKURecord],
    disp_ptr: List[int],
    so_ptr: List[int],
    stock_ptr: List[int],
    logger: PlanogramLogger,
) -> List[Dict[str, Any]]:
    """Allocate one bay and return position-level output rows."""
    rows: List[Dict[str, Any]] = []

    def _row(shelf: int, pos: int, rec: SKURecord) -> Dict[str, Any]:
        return {
            "Store":           store,
            "Bay#":            bay_num,
            "Bay Size":        bay_size,
            "Shelf":           shelf,
            "Position":        pos,
            "SKU":             rec.sku,
            "SKU Type":        rec.sku_type,
            "SKU Description": rec.description,
            "Facing":          rec.facing,
        }

    # Shelf 1 — Display Boards
    for pos, rec in enumerate(
        _consume(disp_pool, disp_ptr, rule.display, store, bay_num, "Display", logger),
        start=1,
    ):
        rows.append(_row(SHELF_DISPLAY, pos, rec))

    # Shelf 2 — Special Order Boards
    for pos, rec in enumerate(
        _consume(so_pool, so_ptr, rule.so, store, bay_num, "SO", logger),
        start=1,
    ):
        rows.append(_row(SHELF_SO, pos, rec))

    # Shelves 3 & 4 — Stock (odd remainder goes to Shelf 3)
    shelf3_count  = math.ceil(rule.stock / 2)
    stock_records = _consume(
        stock_pool, stock_ptr, rule.stock, store, bay_num, "Stock", logger
    )
    for pos, rec in enumerate(stock_records[:shelf3_count], start=1):
        rows.append(_row(SHELF_STOCK_FIRST, pos, rec))
    for pos, rec in enumerate(stock_records[shelf3_count:], start=1):
        rows.append(_row(SHELF_STOCK_SECOND, pos, rec))

    return rows


def allocate_store(
    store: str,
    notes: str,
    bay_list: List[str],
    raw_display: List[SKURecord],
    raw_so: List[SKURecord],
    raw_stock: List[SKURecord],
    logger: PlanogramLogger,
) -> List[Dict[str, Any]]:
    """Run the complete allocation pipeline for one store."""
    # Step 1: Apply Notes rules via stable partition (on unexpanded lists)
    disp_pool  = expand_facing(apply_notes_rules(raw_display, notes))
    so_pool    = expand_facing(apply_notes_rules(raw_so,      notes))
    stock_pool = expand_facing(apply_notes_rules(raw_stock,   notes))

    # Step 2: Independent per-type pointers
    disp_ptr  = [0]
    so_ptr    = [0]
    stock_ptr = [0]
    all_rows: List[Dict[str, Any]] = []

    # Step 3: Allocate each bay in sequence
    for bay_num, bay_size in enumerate(bay_list, start=1):
        rule = BAY_RULES.get(bay_size)
        if rule is None:
            logger.warning(
                f"Unsupported Bay Size '{bay_size}' in Bay {bay_num} — skipping.",
                store=store,
                bay_num=bay_num,
            )
            continue
        all_rows.extend(
            allocate_bay(
                store=store,
                bay_num=bay_num,
                bay_size=bay_size,
                rule=rule,
                disp_pool=disp_pool,
                so_pool=so_pool,
                stock_pool=stock_pool,
                disp_ptr=disp_ptr,
                so_ptr=so_ptr,
                stock_ptr=stock_ptr,
                logger=logger,
            )
        )

    # Step 4: Log any leftover SKUs (diagnostic only)
    for label, pool, ptr in [
        ("Display", disp_pool, disp_ptr),
        ("SO",      so_pool,   so_ptr),
        ("Stock",   stock_pool, stock_ptr),
    ]:
        remaining = len(pool) - ptr[0]
        if remaining > 0:
            logger.warning(
                f"{remaining} {label} position(s) unused after all bays allocated.",
                store=store,
            )

    return all_rows


# ---------------------------------------------------------------------------
# Orchestrator
# ---------------------------------------------------------------------------

def generate_planogram(
    wb_data: WorkbookData,
    logger: PlanogramLogger,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Run the complete planogram generation pipeline.

    Parameters
    ----------
    wb_data:  Parsed workbook data from load_workbook_from_bytes().
    logger:   Fresh PlanogramLogger instance.

    Returns
    -------
    (planogram_df, validation_df)
        planogram_df  — position-level layout for all stores.
        validation_df — all WARNING/ERROR issues logged during generation.
    """
    # Build SKU indexes for all three SKU types
    disp_index  = build_display_index(wb_data.stock_display, wb_data.cols_sd, logger)
    stock_index = build_stock_index(wb_data.stock_display,   wb_data.cols_sd, logger)
    so_index    = build_so_index(wb_data.special_orders,     wb_data.cols_so, logger)

    # Warn on duplicate stores in Store List
    store_col    = wb_data.cols_sl["store"]
    store_series = wb_data.store_list[store_col].apply(_clean_store_id)
    for dup in store_series[store_series.duplicated(keep=False)].unique():
        if dup:
            logger.warning(f"Duplicate store '{dup}' in Store List — all rows processed.")

    all_output_rows: List[Dict[str, Any]] = []
    seen_stores: set = set()

    for _, row in wb_data.store_list.iterrows():
        store = _clean_store_id(row[wb_data.cols_sl["store"]])
        if not store:
            logger.warning("Store column is empty for a row — skipping.")
            continue

        pog_raw = _str(row[wb_data.cols_sl["pog"]])
        lft_raw = _str(row[wb_data.cols_sl["lft"]])
        notes   = _str(row[wb_data.cols_sl["notes"]])

        if store in seen_stores:
            logger.warning(f"Store {store} appears more than once — processing duplicate.", store=store)
        seen_stores.add(store)

        # Parse POG
        pog_bays = parse_pog_string(pog_raw, store=store, logger=logger)
        if not pog_bays:
            logger.warning(
                f"Malformed/empty POG '{pog_raw}' — skipping store.", store=store
            )
            continue

        # Parse LFT (if present)
        lft_bays: List[str] = []
        if lft_raw.lower() not in LFT_IGNORE_VALUES:
            lft_bays = parse_pog_string(lft_raw, store=store, logger=logger)
            if not lft_bays:
                logger.warning(
                    f"Malformed LFT '{lft_raw}' — ignoring LFT.", store=store
                )

        # ── LFT BAYS FIRST, then POG bays ──────────────────────────────────
        bay_list: List[str] = lft_bays + pog_bays

        # Fetch SKU lists for this store
        raw_display = disp_index.get(store, [])
        raw_stock   = stock_index.get(store, [])
        raw_so      = so_index.get(store, [])

        if not raw_display:
            logger.warning(f"No Display SKUs found for store {store}.", store=store)
        if not raw_so:
            logger.warning(f"No SO SKUs found for store {store}.", store=store)
        if not raw_stock:
            logger.warning(f"No Stock SKUs found for store {store}.", store=store)

        try:
            store_rows = allocate_store(
                store=store,
                notes=notes,
                bay_list=bay_list,
                raw_display=raw_display,
                raw_so=raw_so,
                raw_stock=raw_stock,
                logger=logger,
            )
        except Exception as exc:
            logger.error(
                f"Unexpected error allocating store {store}: {exc}", store=store
            )
            continue

        all_output_rows.extend(store_rows)

    planogram_df = pd.DataFrame(all_output_rows, columns=PLANOGRAM_COLUMNS)

    validation_df = (
        pd.DataFrame(logger.issues)
        if logger.issues
        else pd.DataFrame(columns=["Level", "Store", "Bay#", "Message"])
    )

    return planogram_df, validation_df
