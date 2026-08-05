"""
Home Depot Planogram Generator & Auditor — Single-File Streamlit App
=====================================================================
All modules are bundled inline so this file can be deployed directly to
Streamlit Cloud by replacing app.py in your GitHub repository.

Features:
 1. Mode A: Generate Planogram from raw workbook + Automatically run Audit verification.
 2. Mode B: Audit existing Planogram workbook directly.
 3. Appends an 'Audit Report' sheet directly into the downloadable Excel output.
"""
from __future__ import annotations

import math
import re
import traceback
from collections import defaultdict
from dataclasses import dataclass, field
from io import BytesIO
from typing import Any, Dict, List, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook as _openpyxl_load
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="HD Planogram Generator & Auditor",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Aesthetics
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #0f0f0f 0%, #1a1a1a 50%, #111 100%);
    color: #f0f0f0;
}
[data-testid="stHeader"] { background: transparent; }
.hero-banner {
    background: linear-gradient(135deg, #F96302 0%, #cc4f00 60%, #1a1a1a 100%);
    border-radius: 16px; padding: 2rem 2rem 1.5rem; margin-bottom: 1.5rem;
}
.hero-title { font-size: 2.2rem; font-weight: 800; color: #fff; margin: 0 0 0.4rem 0; line-height: 1.1; }
.hero-sub   { font-size: 1rem; color: rgba(255,255,255,0.85); margin: 0; }
.info-card {
    background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1);
    border-radius: 12px; padding: 1.1rem 1.4rem; margin-bottom: 1.2rem;
    font-size: 0.9rem; color: #aaa; line-height: 1.7;
}
.info-card code { background: rgba(249,99,2,0.18); color: #F96302; border-radius: 4px; padding: 1px 6px; }
.stat-card {
    background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.1);
    border-radius: 12px; padding: 1.2rem 1.5rem; border-left: 4px solid #F96302; text-align: center;
}
.stat-value { font-size: 2rem; font-weight: 800; color: #F96302; line-height: 1; }
.stat-label { font-size: 0.8rem; color: #888; margin-top: 0.4rem; }
.section-hdr {
    font-size: 1.05rem; font-weight: 700; color: #f0f0f0;
    border-bottom: 2px solid #F96302; padding-bottom: 0.45rem; margin: 1.8rem 0 1rem;
}
[data-testid="stFileUploader"] section {
    border: 2px dashed #F96302 !important; border-radius: 12px !important;
    background: rgba(249,99,2,0.06) !important;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #F96302, #cc4f00) !important;
    color: #fff !important; font-weight: 700 !important; border: none !important;
    border-radius: 10px !important; padding: 0.65rem 2.2rem !important;
    font-size: 1rem !important; box-shadow: 0 4px 15px rgba(249,99,2,.35) !important;
}
.stDownloadButton > button {
    background: linear-gradient(135deg, #28a745, #1e7e34) !important;
    color: #fff !important; font-weight: 700 !important; border: none !important;
    border-radius: 10px !important; box-shadow: 0 4px 15px rgba(40,167,69,.3) !important;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — MODELS & CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════

@dataclass(frozen=True)
class BayRule:
    display: int
    so:      int
    stock:   int

    @property
    def total(self) -> int:
        return self.display + self.so + self.stock


@dataclass
class SKURecord:
    sku:         str
    description: str
    facing:      int
    sku_type:    str   # "Display" | "SO" | "Stock"


BAY_RULES: Dict[str, BayRule] = {
    "99":  BayRule(display=6, so=2, stock=6),
    "99C": BayRule(display=4, so=2, stock=4),
    "87":  BayRule(display=4, so=1, stock=4),
    "75":  BayRule(display=4, so=1, stock=4),
    "51":  BayRule(display=2, so=1, stock=2),
}

NOTES_RULES: List[Dict[str, Any]] = [
    {
        "trigger_keywords": ["baja", "vigo"],
        "sku_keywords":     ["baja", "vigo"],
    },
]

SHEET_STORE_LIST    = "Store List"
SHEET_STOCK_DISPLAY = "Stock SKUs and Displays"
SHEET_SO            = "Special Order Boards"

OUTPUT_SHEET_PLANOGRAM  = "Generated Planogram"
OUTPUT_SHEET_VALIDATION = "Validation"
OUTPUT_SHEET_AUDIT      = "Audit Report"

PLANOGRAM_COLUMNS: List[str] = [
    "Store", "Bay#", "Bay Size", "Shelf", "Position",
    "SKU", "SKU Type", "SKU Description", "Facing",
]

SHELF_DISPLAY      = 1
SHELF_SO           = 2
SHELF_STOCK_FIRST  = 3
SHELF_STOCK_SECOND = 4

LFT_IGNORE_VALUES: frozenset = frozenset(["-", "", "none", "null", "n/a", "na"])
VALID_FACINGS: frozenset = frozenset([1, 2])

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — LOGGER & PARSER
# ══════════════════════════════════════════════════════════════════════════════

class PlanogramLogger:
    def __init__(self) -> None:
        self.issues: List[Dict[str, Any]] = []
        self._seen: set = set()

    def _append(self, level: str, msg: str, store: Any = "", bay_num: Any = "") -> None:
        key = (level, str(store), str(bay_num), msg)
        if key in self._seen:
            return
        self._seen.add(key)
        self.issues.append({"Level": level, "Store": str(store), "Bay#": str(bay_num), "Message": msg})

    def warning(self, msg: str, store: Any = "", bay_num: Any = "") -> None:
        self._append("WARNING", msg, store, bay_num)

    def error(self, msg: str, store: Any = "", bay_num: Any = "") -> None:
        self._append("ERROR", msg, store, bay_num)

    def info(self, msg: str) -> None:
        pass


_RE_LEADING_COUNT = re.compile(r"^\s*(\d+)\s+bay\s*[-\u2013]?\s*", re.IGNORECASE)
_RE_VALID_TOKEN = re.compile(r"^[A-Za-z0-9]+(?:-[A-Za-z0-9]+)?$")
_RE_MULTIPLIER_TOKEN = re.compile(r"^(\d+)-([A-Za-z0-9]+)$")
_RE_BAYSIZE_PREFIX = re.compile(r"^(\d+[A-Za-z]{0,2}(?:-\d+[A-Za-z]{0,2})?)")


def parse_pog_string(pog: str, store: Any = "", logger: PlanogramLogger | None = None) -> List[str]:
    if not isinstance(pog, str) or not pog.strip():
        return []

    pog = pog.strip()
    m_count = _RE_LEADING_COUNT.match(pog)
    if not m_count:
        if logger:
            logger.warning(f"Cannot parse bay count from POG: '{pog}'", store=store)
        return []

    declared_count: int = int(m_count.group(1))
    remainder: str = pog[m_count.end():].strip()

    raw_tokens = [t.strip() for t in remainder.split(",")]
    bay_tokens_raw: List[str] = []

    for tok in raw_tokens:
        if _RE_VALID_TOKEN.match(tok):
            bay_tokens_raw.append(tok)
        else:
            m_prefix = _RE_BAYSIZE_PREFIX.match(tok)
            if m_prefix:
                bay_tokens_raw.append(m_prefix.group(1))
            break

    if not bay_tokens_raw:
        if logger:
            logger.warning(f"No bay tokens found in POG: '{pog}'", store=store)
        return []

    all_multiplier = all(_RE_MULTIPLIER_TOKEN.match(tok) for tok in bay_tokens_raw)

    if all_multiplier:
        expanded: List[str] = []
        for tok in bay_tokens_raw:
            m = _RE_MULTIPLIER_TOKEN.match(tok)
            expanded.extend([m.group(2).upper()] * int(m.group(1)))
    elif len(bay_tokens_raw) == 1:
        expanded = [bay_tokens_raw[0].upper()] * declared_count
    else:
        expanded = [tok.upper() for tok in bay_tokens_raw]

    if len(expanded) != declared_count and logger:
        logger.warning(
            f"Parsed {len(expanded)} bay(s) but header declares {declared_count}: '{pog}'",
            store=store,
        )

    return expanded

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — WORKBOOK LOADER & ALLOCATOR
# ══════════════════════════════════════════════════════════════════════════════

def _str(val: Any) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def _clean_store_id(val: Any) -> str:
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
    return str(name).strip().lower()


def _find_col(df: pd.DataFrame, candidates: List[str], fallback_pos: int, sheet_name: str, logger: PlanogramLogger | None = None) -> str:
    cols = list(df.columns)
    norm_map = {_norm(c): c for c in cols}

    for cand in candidates:
        if _norm(cand) in norm_map:
            return norm_map[_norm(cand)]

    if fallback_pos < len(cols):
        actual = cols[fallback_pos]
        if logger:
            logger.info(f"[{sheet_name}] Header '{candidates[0]}' not found; using position {fallback_pos} ('{actual}').")
        return actual

    raise ValueError(f"[{sheet_name}] Cannot find column '{candidates[0]}' by name or position {fallback_pos}.")


@dataclass
class WorkbookData:
    store_list:     pd.DataFrame
    stock_display:  pd.DataFrame
    special_orders: pd.DataFrame
    cols_sl: Dict[str, str] = field(default_factory=dict)
    cols_sd: Dict[str, str] = field(default_factory=dict)
    cols_so: Dict[str, str] = field(default_factory=dict)


def load_workbook_from_bytes(file_bytes: bytes, logger: PlanogramLogger | None = None) -> WorkbookData:
    try:
        xl = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    except Exception as exc:
        raise RuntimeError(f"Cannot open workbook: {exc}") from exc

    def _get_sheet(desired: str) -> str:
        for name in xl.sheet_names:
            if name.strip().lower() == desired.strip().lower():
                return name
        raise ValueError(f"Sheet '{desired}' not found. Available: {xl.sheet_names}")

    s_sl = _get_sheet(SHEET_STORE_LIST)
    s_sd = _get_sheet(SHEET_STOCK_DISPLAY)
    s_so = _get_sheet(SHEET_SO)

    df_sl = xl.parse(s_sl, dtype=str)
    df_sd = xl.parse(s_sd, dtype=str)
    df_so = xl.parse(s_so, dtype=str)

    for df in (df_sl, df_sd, df_so):
        df.columns = [str(c).strip() for c in df.columns]
        df.dropna(how="all", inplace=True)
        df.reset_index(drop=True, inplace=True)

    cols_sl = {
        "store": _find_col(df_sl, ["Store"], 0, s_sl, logger),
        "pog":   _find_col(df_sl, ["Current Store POG", "Store POG", "POG"], 1, s_sl, logger),
        "lft":   _find_col(df_sl, ["Current LFT", "LFT"], 2, s_sl, logger),
        "notes": _find_col(df_sl, ["Notes", "Note"], 3, s_sl, logger),
    }

    sd_cols = list(df_sd.columns)
    cols_sd = {
        "store":      _find_col(df_sd, ["Store"], 0, s_sd, logger),
        "stock_sku":  _find_col(df_sd, ["Stock SKU", "Stock Sku"], 1, s_sd, logger),
        "stock_desc": _find_col(df_sd, ["Stock Description", "Stock Desc"], 2, s_sd, logger),
        "stock_face": _find_col(df_sd, ["Facings", "Facing"], 4, s_sd, logger),
        "disp_sku":   _find_col(df_sd, ["Display SKU", "Display Sku"], 5, s_sd, logger),
        "disp_desc":  _find_col(df_sd, ["Display Description", "Display Desc"], 6, s_sd, logger),
        "disp_face":  _find_col(df_sd, ["Display Facing", "Facings.1"], 7 if len(sd_cols) <= 9 else 8, s_sd, logger),
        "cf":         _find_col(df_sd, ["CF"], len(sd_cols) - 1, s_sd, logger),
    }

    so_cols = list(df_so.columns)
    cols_so = {
        "store":   _find_col(df_so, ["Store"], 0, s_so, logger),
        "so_sku":  _find_col(df_so, ["SO SKU", "So Sku", "SKU"], 1, s_so, logger),
        "so_desc": _find_col(df_so, ["Description"], 2, s_so, logger),
        "so_face": _find_col(df_so, ["Facings", "Facing"], 4, s_so, logger),
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


def _parse_facing(raw: Any, store: Any, sku: Any, logger: PlanogramLogger) -> int:
    try:
        val = int(float(str(raw).strip()))
    except (ValueError, TypeError):
        logger.warning(f"Invalid Facing '{raw}' for SKU {sku} — defaulting to 1.", store=store)
        return 1
    if val not in VALID_FACINGS:
        logger.warning(f"Facing '{raw}' for SKU {sku} not in {{1,2}} — defaulting to 1.", store=store)
        return 1
    return val


def build_display_index(df: pd.DataFrame, cols: Dict[str, str], logger: PlanogramLogger) -> Dict[str, List[SKURecord]]:
    index: Dict[str, List[SKURecord]] = defaultdict(list)
    for _, row in df.iterrows():
        store  = _clean_store_id(row[cols["store"]])
        sku    = _str(row[cols["disp_sku"]])
        desc   = _str(row[cols["disp_desc"]])
        disp_f  = _parse_facing(row[cols["disp_face"]], store, sku, logger)
        stock_f = _parse_facing(row[cols["stock_face"]], store, sku, logger)
        facing  = max(disp_f, stock_f)
        if not store or not sku:
            continue
        index[store].append(SKURecord(sku=sku, description=desc, facing=facing, sku_type="Display"))
    return dict(index)


def build_stock_index(df: pd.DataFrame, cols: Dict[str, str], logger: PlanogramLogger) -> Dict[str, List[SKURecord]]:
    index: Dict[str, List[SKURecord]] = defaultdict(list)
    for _, row in df.iterrows():
        store  = _clean_store_id(row[cols["store"]])
        sku    = _str(row[cols["stock_sku"]])
        desc   = _str(row[cols["stock_desc"]])
        facing = _parse_facing(row[cols["stock_face"]], store, sku, logger)
        if not store or not sku:
            continue
        index[store].append(SKURecord(sku=sku, description=desc, facing=facing, sku_type="Stock"))
    return dict(index)


def build_so_index(df: pd.DataFrame, cols: Dict[str, str], logger: PlanogramLogger) -> Dict[str, List[SKURecord]]:
    index: Dict[str, List[SKURecord]] = defaultdict(list)
    for _, row in df.iterrows():
        store  = _clean_store_id(row[cols["store"]])
        sku    = _str(row[cols["so_sku"]])
        desc   = _str(row[cols["so_desc"]])
        facing = _parse_facing(row[cols["so_face"]], store, sku, logger)
        if not store or not sku:
            continue
        index[store].append(SKURecord(sku=sku, description=desc, facing=facing, sku_type="SO"))
    return dict(index)


def expand_facing(records: List[SKURecord]) -> List[SKURecord]:
    expanded: List[SKURecord] = []
    for rec in records:
        for _ in range(max(1, rec.facing)):
            expanded.append(rec)
    return expanded


def apply_notes_rules(records: List[SKURecord], notes: str) -> List[SKURecord]:
    if not notes or not isinstance(notes, str):
        return records

    notes_lower = notes.strip().lower()
    result = list(records)

    for rule in NOTES_RULES:
        triggers = list(rule.get("trigger_keywords", []))
        if not any(tr.lower() in notes_lower for tr in triggers if tr):
            continue

        sku_kws = [kw.lower() for kw in rule.get("sku_keywords", []) if kw]
        if not sku_kws:
            continue

        head: List[SKURecord] = []
        tail: List[SKURecord] = []
        for rec in result:
            if any(kw in rec.description.lower() for kw in sku_kws):
                tail.append(rec)
            else:
                head.append(rec)
        result = head + tail

    return result


def _consume(pool: List[SKURecord], pointer: List[int], count: int, store: Any, bay_num: int, sku_type: str, logger: PlanogramLogger) -> List[SKURecord]:
    start = pointer[0]
    end   = start + count
    taken = pool[start:end]
    if len(taken) < count:
        logger.warning(f"Not enough {sku_type} SKUs for Bay {bay_num}: need {count}, have {len(taken)} remaining.", store=store, bay_num=bay_num)
    pointer[0] = min(end, len(pool))
    return taken


def allocate_bay(store: Any, bay_num: int, bay_size: str, rule: BayRule,
                 disp_pool: List[SKURecord], so_pool: List[SKURecord], stock_pool: List[SKURecord],
                 disp_ptr: List[int], so_ptr: List[int], stock_ptr: List[int], logger: PlanogramLogger) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    def _row(shelf: int, pos: int, rec: SKURecord) -> Dict[str, Any]:
        return {
            "Store": store, "Bay#": bay_num, "Bay Size": bay_size,
            "Shelf": shelf, "Position": pos,
            "SKU": rec.sku, "SKU Type": rec.sku_type,
            "SKU Description": rec.description, "Facing": rec.facing,
        }

    for pos, rec in enumerate(_consume(disp_pool, disp_ptr, rule.display, store, bay_num, "Display", logger), start=1):
        rows.append(_row(SHELF_DISPLAY, pos, rec))

    for pos, rec in enumerate(_consume(so_pool, so_ptr, rule.so, store, bay_num, "SO", logger), start=1):
        rows.append(_row(SHELF_SO, pos, rec))

    shelf3_count  = math.ceil(rule.stock / 2)
    stock_records = _consume(stock_pool, stock_ptr, rule.stock, store, bay_num, "Stock", logger)
    for pos, rec in enumerate(stock_records[:shelf3_count], start=1):
        rows.append(_row(SHELF_STOCK_FIRST, pos, rec))
    for pos, rec in enumerate(stock_records[shelf3_count:], start=1):
        rows.append(_row(SHELF_STOCK_SECOND, pos, rec))

    return rows


def allocate_store(store: str, notes: str, bay_list: List[str], raw_display: List[SKURecord], raw_so: List[SKURecord], raw_stock: List[SKURecord], logger: PlanogramLogger) -> List[Dict[str, Any]]:
    disp_pool  = expand_facing(apply_notes_rules(raw_display, notes))
    so_pool    = expand_facing(apply_notes_rules(raw_so,      notes))
    stock_pool = expand_facing(apply_notes_rules(raw_stock,   notes))

    disp_ptr, so_ptr, stock_ptr = [0], [0], [0]
    all_rows: List[Dict[str, Any]] = []

    for bay_num, bay_size in enumerate(bay_list, start=1):
        rule = BAY_RULES.get(bay_size)
        if rule is None:
            logger.warning(f"Unsupported Bay Size '{bay_size}' in Bay {bay_num} — skipping.", store=store, bay_num=bay_num)
            continue
        all_rows.extend(allocate_bay(store=store, bay_num=bay_num, bay_size=bay_size, rule=rule, disp_pool=disp_pool, so_pool=so_pool, stock_pool=stock_pool, disp_ptr=disp_ptr, so_ptr=so_ptr, stock_ptr=stock_ptr, logger=logger))

    for label, pool, ptr in [("Display", disp_pool, disp_ptr), ("SO", so_pool, so_ptr), ("Stock", stock_pool, stock_ptr)]:
        remaining = len(pool) - ptr[0]
        if remaining > 0:
            logger.warning(f"{remaining} {label} position(s) unused after all bays allocated.", store=store)

    return all_rows


def generate_planogram(wb_data: WorkbookData, logger: PlanogramLogger) -> Tuple[pd.DataFrame, pd.DataFrame]:
    disp_index  = build_display_index(wb_data.stock_display, wb_data.cols_sd, logger)
    stock_index = build_stock_index(wb_data.stock_display,   wb_data.cols_sd, logger)
    so_index    = build_so_index(wb_data.special_orders,     wb_data.cols_so, logger)

    all_output_rows: List[Dict[str, Any]] = []
    seen_stores: set = set()

    for _, row in wb_data.store_list.iterrows():
        store = _clean_store_id(row[wb_data.cols_sl["store"]])
        if not store:
            continue

        pog_raw = _str(row[wb_data.cols_sl["pog"]])
        lft_raw = _str(row[wb_data.cols_sl["lft"]])
        notes   = _str(row[wb_data.cols_sl["notes"]])

        seen_stores.add(store)

        pog_bays = parse_pog_string(pog_raw, store=store, logger=logger)
        if not pog_bays:
            continue

        lft_bays: List[str] = []
        if lft_raw.lower() not in LFT_IGNORE_VALUES:
            lft_bays = parse_pog_string(lft_raw, store=store, logger=logger)

        bay_list: List[str] = lft_bays + pog_bays

        raw_display = disp_index.get(store, [])
        raw_stock   = stock_index.get(store, [])
        raw_so      = so_index.get(store, [])

        store_rows = allocate_store(store=store, notes=notes, bay_list=bay_list, raw_display=raw_display, raw_so=raw_so, raw_stock=raw_stock, logger=logger)
        all_output_rows.extend(store_rows)

    planogram_df = pd.DataFrame(all_output_rows, columns=PLANOGRAM_COLUMNS)
    validation_df = pd.DataFrame(logger.issues) if logger.issues else pd.DataFrame(columns=["Level", "Store", "Bay#", "Message"])
    return planogram_df, validation_df

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — AUDIT ENGINE
# ══════════════════════════════════════════════════════════════════════════════

def audit_planogram_df(df_stores: pd.DataFrame, df_pog: pd.DataFrame, notes_filter: str = "baja|vigo", keywords: List[str] = ["BAJA", "VIGO"], categories: List[str] = ["Display", "SO", "Stock"]) -> pd.DataFrame:
    if 'Store' not in df_stores.columns or 'Notes' not in df_stores.columns:
        return pd.DataFrame()

    filtered_stores_df = df_stores[df_stores['Notes'].astype(str).str.contains(notes_filter, case=False, na=False)]
    filtered_stores = filtered_stores_df['Store'].astype(str).unique().tolist()

    kw_pattern = "|".join([k.strip() for k in keywords if k.strip()])
    results = []

    for store_id in filtered_stores:
        st_data = df_pog[df_pog['Store'].astype(str) == str(store_id)]
        if len(st_data) == 0:
            results.append({'Store': store_id, 'Status': 'MISSING', 'Passed': False, 'Details': 'Store missing in Planogram sheet'})
            continue

        store_pass = True
        cat_statuses = []

        for c in categories:
            sub = st_data[st_data['SKU Type'] == c]
            bv = sub[sub['SKU Description'].astype(str).str.contains(kw_pattern, case=False, na=False)]
            non_bv = sub[~sub['SKU Description'].astype(str).str.contains(kw_pattern, case=False, na=False)]

            bv_idx = bv.index.tolist()
            non_bv_idx = non_bv.index.tolist()

            if len(bv_idx) == 0:
                c_status = 'NO_TARGET_SKUS'
            elif len(non_bv_idx) == 0:
                c_status = 'ALL_TARGET_SKUS'
            else:
                if min(bv_idx) > max(non_bv_idx):
                    c_status = 'PASS'
                else:
                    c_status = 'FAIL'
                    store_pass = False
            cat_statuses.append(f"{c}: {c_status} ({len(bv)} SKUs)")

        results.append({
            'Store': store_id,
            'Status': 'PASS' if store_pass else 'FAIL',
            'Passed': store_pass,
            'Details': " | ".join(cat_statuses)
        })

    return pd.DataFrame(results)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — EXCEL WRITER (WITH AUDIT REPORT SHEET)
# ══════════════════════════════════════════════════════════════════════════════

_HEADER_FONT  = Font(bold=True, color="FFFFFF")
_HD_ORANGE    = PatternFill(fill_type="solid", fgColor="F96302")
_ERROR_RED    = PatternFill(fill_type="solid", fgColor="CC0000")
_NAVY_BLUE    = PatternFill(fill_type="solid", fgColor="1E293B")
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _style_header(ws, fill: PatternFill = _HD_ORANGE) -> None:
    for cell in ws[1]:
        cell.font      = _HEADER_FONT
        cell.fill      = fill
        cell.alignment = _HEADER_ALIGN
    ws.row_dimensions[1].height = 22


def _auto_width(ws, max_width: int = 60) -> None:
    for col_cells in ws.columns:
        length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col_cells)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(length + 3, max_width)


def write_output_bytes(input_bytes: bytes, planogram_df: pd.DataFrame, validation_df: pd.DataFrame, audit_df: pd.DataFrame = None) -> bytes:
    wb = _openpyxl_load(BytesIO(input_bytes))

    for name in [OUTPUT_SHEET_PLANOGRAM, OUTPUT_SHEET_VALIDATION, OUTPUT_SHEET_AUDIT]:
        if name in wb.sheetnames:
            del wb[name]

    ws_pog = wb.create_sheet(title=OUTPUT_SHEET_PLANOGRAM)
    ws_pog.append(PLANOGRAM_COLUMNS)
    for _, row in planogram_df.iterrows():
        ws_pog.append([row[col] for col in PLANOGRAM_COLUMNS])
    _style_header(ws_pog, fill=_HD_ORANGE)
    _auto_width(ws_pog)
    ws_pog.freeze_panes = "A2"

    val_cols: List[str] = ["Level", "Store", "Bay#", "Message"]
    ws_val = wb.create_sheet(title=OUTPUT_SHEET_VALIDATION)
    ws_val.append(val_cols)
    for _, row in validation_df.iterrows():
        ws_val.append([_str(row.get(c, "")) for c in val_cols])
    _style_header(ws_val, fill=_ERROR_RED)
    _auto_width(ws_val)
    ws_val.freeze_panes = "A2"

    if audit_df is not None and not audit_df.empty:
        ws_aud = wb.create_sheet(title=OUTPUT_SHEET_AUDIT)
        aud_cols = ["Store", "Status", "Details"]
        ws_aud.append(aud_cols)
        for _, row in audit_df.iterrows():
            ws_aud.append([_str(row.get(c, "")) for c in aud_cols])
        _style_header(ws_aud, fill=_NAVY_BLUE)
        _auto_width(ws_aud)
        ws_aud.freeze_panes = "A2"

    output_bio = BytesIO()
    wb.save(output_bio)
    output_bio.seek(0)
    return output_bio.read()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — STREAMLIT UI & WORKFLOWS
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="hero-banner">
  <div class="hero-title">🏗️ Home Depot Planogram Generator & Auditor</div>
  <p class="hero-sub">Generate shelf layouts & perform instant compliance verification for forced SKU end-position rules.</p>
</div>
""", unsafe_allow_html=True)

st.sidebar.header("🎯 Workflow Selection")
app_mode = st.sidebar.radio(
    "Choose Action",
    ["Generate & Audit Planogram (Full Flow)", "Audit Existing Planogram Only"],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ Audit Rule Settings")
notes_filter = st.sidebar.text_input("Store Notes Filter (Regex)", "baja|vigo")
keywords_str = st.sidebar.text_input("Target SKU Keywords", "BAJA, VIGO")
keywords = [k.strip() for k in keywords_str.split(",") if k.strip()]
categories = ["Display", "SO", "Stock"]

# ==============================================================================
# MODE A: GENERATE & AUDIT PLANOGRAM
# ==============================================================================
if app_mode == "Generate & Audit Planogram (Full Flow)":
    st.markdown("""
    <div class="info-card">
      <strong>Full Automation Flow:</strong><br>
      1. Upload raw Home Depot workbook containing <code>Store List</code>, <code>Stock SKUs and Displays</code>, and <code>Special Order Boards</code>.<br>
      2. App allocates SKUs into <code>Generated Planogram</code>.<br>
      3. App automatically audits all forced stores (e.g., "Force Baja & Vigos to the end") and displays an interactive compliance dashboard.<br>
      4. Download the final Excel file with <code>Generated Planogram</code>, <code>Validation</code>, and <code>Audit Report</code> sheets bundled together!
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Upload Raw Excel Workbook (.xlsx)", type=["xlsx", "xls"])
    run_clicked = st.button("▶ Generate Planogram & Run Audit", type="primary")

    if run_clicked:
        if uploaded_file is None:
            st.error("⚠️ Please upload an Excel workbook first.")
            st.stop()

        file_bytes = uploaded_file.getvalue()
        logger     = PlanogramLogger()

        try:
            with st.spinner("📂 Loading workbook & parsing sheets…"):
                wb_data = load_workbook_from_bytes(file_bytes, logger)

            with st.spinner("⚙️ Allocating SKUs and generating planogram layout…"):
                planogram_df, validation_df = generate_planogram(wb_data, logger)

            with st.spinner("🔍 Running automated audit verification on generated planogram…"):
                audit_df = audit_planogram_df(wb_data.store_list, planogram_df, notes_filter, keywords, categories)

            with st.spinner("💾 Bundling Planogram + Audit Report into final Excel output…"):
                output_bytes = write_output_bytes(file_bytes, planogram_df, validation_df, audit_df)

        except Exception as exc:
            st.error(f"❌ Processing Error: {exc}")
            with st.expander("Technical Error Details"):
                st.code(traceback.format_exc())
            st.stop()

        st.markdown("---")
        st.markdown('<div class="section-hdr">🎉 Generation & Audit Complete</div>', unsafe_allow_html=True)

        n_stores = planogram_df["Store"].nunique() if not planogram_df.empty else 0
        n_rows   = len(planogram_df)
        n_issues = len(validation_df)

        filtered_count = len(audit_df)
        pass_count = int(audit_df['Passed'].sum()) if len(audit_df) > 0 else 0
        pass_rate = (pass_count / filtered_count * 100) if filtered_count > 0 else 0.0

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f'<div class="stat-card"><div class="stat-value">{n_stores:,}</div><div class="stat-label">Stores Processed</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="stat-card"><div class="stat-value">{n_rows:,}</div><div class="stat-label">Planogram Rows</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="stat-card"><div class="stat-value" style="color:#34d399;">{pass_count}/{filtered_count}</div><div class="stat-label">Audited Stores Passed</div></div>', unsafe_allow_html=True)
        with c4:
            st.markdown(f'<div class="stat-card"><div class="stat-value" style="color:#34d399;">{pass_rate:.1f}%</div><div class="stat-label">Compliance Rate</div></div>', unsafe_allow_html=True)

        st.markdown("")
        st.download_button(
            label="📥 Download Complete Planogram Output with Audit Report (.xlsx)",
            data=output_bytes,
            file_name=f"Planogram_Output_Audited.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("---")
        tab1, tab2, tab3 = st.tabs(["📋 Store Audit Report", "👁️ Planogram Preview", "⚠️ Validation Issues"])

        with tab1:
            st.subheader("Audit Results Breakdown")
            st.dataframe(audit_df[['Store', 'Status', 'Details']], use_container_width=True, hide_index=True)

        with tab2:
            st.subheader("Generated Planogram Preview")
            st.dataframe(planogram_df.head(200), use_container_width=True, hide_index=True)

        with tab3:
            st.subheader("Validation Log")
            if n_issues > 0:
                st.dataframe(validation_df, use_container_width=True, hide_index=True)
            else:
                st.success("✅ No validation issues found.")

# ==============================================================================
# MODE B: AUDIT EXISTING PLANOGRAM ONLY
# ==============================================================================
else:
    st.markdown("""
    <div class="info-card">
      <strong>Standalone Audit Mode:</strong><br>
      Upload an existing Planogram workbook containing both <code>Store List</code> and <code>Generated Planogram</code> sheets to audit keyword SKU placement.
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Upload Existing Planogram Workbook (.xlsx)", type=["xlsx", "xls"])
    if uploaded_file is not None:
        xl = pd.ExcelFile(uploaded_file)
        store_sheet = st.selectbox("Store Metadata Sheet", xl.sheet_names, index=0 if "Store List" in xl.sheet_names else 0)
        pog_sheet = st.selectbox("Generated Planogram Sheet", xl.sheet_names, index=xl.sheet_names.index("Generated Planogram") if "Generated Planogram" in xl.sheet_names else min(1, len(xl.sheet_names)-1))

        if st.button("▶ Run Audit on File", type="primary"):
            df_s = pd.read_excel(uploaded_file, sheet_name=store_sheet)
            df_p = pd.read_excel(uploaded_file, sheet_name=pog_sheet)

            audit_df = audit_planogram_df(df_s, df_p, notes_filter, keywords, categories)

            st.markdown("---")
            st.markdown('<div class="section-hdr">📊 Audit Results</div>', unsafe_allow_html=True)

            filtered_count = len(audit_df)
            pass_count = int(audit_df['Passed'].sum()) if len(audit_df) > 0 else 0
            fail_count = filtered_count - pass_count
            pass_rate = (pass_count / filtered_count * 100) if filtered_count > 0 else 0.0

            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f'<div class="stat-card"><div class="stat-value">{filtered_count}</div><div class="stat-label">Stores Filtered</div></div>', unsafe_allow_html=True)
            with c2:
                st.markdown(f'<div class="stat-card"><div class="stat-value" style="color:#34d399;">{pass_count}</div><div class="stat-label">Passed</div></div>', unsafe_allow_html=True)
            with c3:
                st.markdown(f'<div class="stat-card"><div class="stat-value" style="color:#34d399;">{pass_rate:.1f}%</div><div class="stat-label">Compliance Rate</div></div>', unsafe_allow_html=True)

            st.markdown("")
            st.dataframe(audit_df[['Store', 'Status', 'Details']], use_container_width=True, hide_index=True)

            csv_data = audit_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📄 Download Audit Results CSV",
                data=csv_data,
                file_name="Planogram_Audit_Results.csv",
                mime="text/csv",
            )

st.markdown("---")
st.markdown('<p style="text-align:center;color:#666;font-size:.8rem;">Home Depot Planogram Generator & Auditor</p>', unsafe_allow_html=True)
