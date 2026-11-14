"""
HD Planogram Generator and Auditor: Single-File Streamlit App
=============================================================
All modules and ChaiUI presentation components are bundled inline so this
file can be deployed directly to Streamlit Cloud or run locally.

Features:
 1. Generate planogram: Allocates bays from workbook and runs audit verification.
 2. Audit existing planogram: Audits keyword SKU placement directly from uploaded planogram.
 3. Store SKU generator: Generates store SKU mapping sheet from SKU and Store sheets.
"""
from __future__ import annotations

import math
import re
import time
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

# ==============================================================================
# SECTION 1: PAGE CONFIG AND LOADING SCREEN
# ==============================================================================
st.set_page_config(
    page_title="HD Planogram Generator",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 2-3 second animated loading screen on launch
if "app_initialized" not in st.session_state:
    loading_placeholder = st.empty()
    with loading_placeholder.container():
        st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700&family=Montserrat:wght@400;500;600&display=swap');
        .splash-screen {
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            background-color: #0a0a0a;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            z-index: 9999999;
            animation: splashAnim 2.4s cubic-bezier(0.4, 0, 0.2, 1) forwards;
            pointer-events: none;
        }
        @keyframes splashAnim {
            0% {
                opacity: 0;
                transform: scale(0.98);
            }
            20% {
                opacity: 1;
                transform: scale(1);
            }
            80% {
                opacity: 1;
                transform: scale(1);
            }
            100% {
                opacity: 0;
                transform: scale(1.02);
                visibility: hidden;
            }
        }
        .splash-title {
            font-family: 'Manrope', sans-serif;
            font-size: 2.5rem;
            font-weight: 700;
            color: #fafafa;
            letter-spacing: -0.03em;
            margin-bottom: 0.5rem;
            text-align: center;
        }
        .splash-divider {
            width: 48px;
            height: 1px;
            background: rgba(255, 255, 255, 0.15);
            margin: 1.25rem auto;
        }
        .splash-subtitle {
            font-family: 'Manrope', sans-serif;
            font-size: 1rem;
            color: #a3a3a3;
            text-align: center;
        }
        .splash-highlight {
            font-family: 'Montserrat', sans-serif;
            font-weight: 500;
            color: #ffedd5;
        }
        </style>
        <div class="splash-screen">
            <div class="splash-title">HD Planogram Generator</div>
            <div class="splash-divider"></div>
            <div class="splash-subtitle">Automated bay layouts with <span class="splash-highlight">compliance audit</span> verification</div>
        </div>
        """, unsafe_allow_html=True)
        time.sleep(2.4)
    st.session_state["app_initialized"] = True
    loading_placeholder.empty()
    st.rerun()

# ==============================================================================
# SECTION 2: CHAIUI DESIGN SYSTEM AND STYLING
# ==============================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@300;400;500;600;700&family=Montserrat:wght@400;500;600;700&display=swap');

/* Base page resets and dark palette */
html, body, [class*="css"], [data-testid="stMarkdownContainer"] p, [data-testid="stWidgetLabel"] {
    font-family: 'Manrope', -apple-system, BlinkMacSystemFont, sans-serif !important;
}

[data-testid="stAppViewContainer"] {
    background-color: #0a0a0a !important;
    color: #fafafa !important;
}

[data-testid="stHeader"] {
    background-color: transparent !important;
}

[data-testid="stSidebar"] {
    background-color: #121212 !important;
    border-right: 1px solid rgba(255, 255, 255, 0.10) !important;
}

[data-testid="stSidebar"] * {
    color: #fafafa !important;
}

/* Headings */
h1, h2, h3, h4, h5, h6 {
    font-family: 'Manrope', sans-serif !important;
    font-weight: 600 !important;
    color: #fafafa !important;
    letter-spacing: -0.02em !important;
}

/* ChaiUI Hero Banner */
.chai-hero {
    background: rgba(255, 255, 255, 0.02);
    border: 1px solid rgba(255, 255, 255, 0.10);
    border-radius: 12px;
    padding: 2.2rem 2.4rem;
    margin-bottom: 1.5rem;
    box-shadow: none !important;
}

.chai-hero-title {
    font-family: 'Manrope', sans-serif;
    font-size: 2.1rem;
    font-weight: 700;
    color: #fafafa;
    margin: 0 0 0.5rem 0;
    line-height: 1.2;
    letter-spacing: -0.03em;
}

.chai-hero-sub {
    font-family: 'Manrope', sans-serif;
    font-size: 0.98rem;
    color: #a3a3a3;
    margin: 0;
    line-height: 1.6;
}

/* Highlight phrase rule: one per paragraph, Montserrat, highlight cream */
.highlight {
    font-family: 'Montserrat', sans-serif;
    font-weight: 500;
    color: #ffedd5;
}

/* ChaiUI Card */
.chai-card {
    background: rgba(255, 255, 255, 0.02);
    border: 1px solid rgba(255, 255, 255, 0.10);
    border-radius: 12px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 1.25rem;
    box-shadow: none !important;
    color: #a3a3a3;
    font-size: 0.92rem;
    line-height: 1.7;
}

.chai-card strong {
    color: #fafafa;
    font-family: 'Montserrat', sans-serif;
    font-weight: 600;
}

.chai-card code {
    background: rgba(255, 255, 255, 0.06);
    color: #ffedd5;
    border-radius: 4px;
    padding: 2px 6px;
    font-size: 0.85rem;
}

/* Stat Cards */
.chai-stat-card {
    background: rgba(255, 255, 255, 0.02);
    border: 1px solid rgba(255, 255, 255, 0.10);
    border-radius: 12px;
    padding: 1.2rem 1.4rem;
    text-align: center;
    box-shadow: none !important;
}

.chai-stat-value {
    font-family: 'Montserrat', sans-serif;
    font-size: 2rem;
    font-weight: 600;
    color: #fafafa;
    line-height: 1.1;
}

.chai-stat-value.highlight-green {
    color: #4ade80;
}

.chai-stat-label {
    font-family: 'Manrope', sans-serif;
    font-size: 0.82rem;
    color: #a3a3a3;
    margin-top: 0.4rem;
}

/* Section Header */
.chai-section-hdr {
    font-family: 'Manrope', sans-serif;
    font-size: 1.1rem;
    font-weight: 600;
    color: #fafafa;
    padding-bottom: 0.5rem;
    margin: 1.6rem 0 1rem;
    border-bottom: 1px solid rgba(255, 255, 255, 0.10);
}

/* Asymmetric corner buttons */
.stButton > button[kind="primary"], button[kind="primary"] {
    background: #fafafa !important;
    color: #0a0a0a !important;
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    border: 1px solid #fafafa !important;
    border-radius: 0 8px 0 8px !important;
    padding: 0.65rem 2rem !important;
    box-shadow: none !important;
    transition: background 0.2s ease, opacity 0.2s ease !important;
}

.stButton > button[kind="primary"]:hover, button[kind="primary"]:hover {
    background: #e5e5e5 !important;
    color: #0a0a0a !important;
    border-color: #e5e5e5 !important;
    box-shadow: none !important;
}

.stButton > button[kind="secondary"], button[kind="secondary"], .stDownloadButton > button {
    background: transparent !important;
    color: #fafafa !important;
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 500 !important;
    font-size: 0.95rem !important;
    border: 1px solid rgba(255, 255, 255, 0.18) !important;
    border-radius: 8px 0 8px 0 !important;
    padding: 0.65rem 2rem !important;
    box-shadow: none !important;
    transition: background 0.2s ease, border-color 0.2s ease !important;
}

.stButton > button[kind="secondary"]:hover, button[kind="secondary"]:hover, .stDownloadButton > button:hover {
    background: rgba(255, 255, 255, 0.05) !important;
    border-color: rgba(255, 255, 255, 0.35) !important;
    color: #fafafa !important;
    box-shadow: none !important;
}

/* File Uploader styling */
[data-testid="stFileUploader"] section {
    border: 1px dashed rgba(255, 255, 255, 0.15) !important;
    border-radius: 12px !important;
    background: rgba(255, 255, 255, 0.02) !important;
    box-shadow: none !important;
}

[data-testid="stFileUploader"] section:hover {
    border-color: rgba(255, 255, 255, 0.30) !important;
}

[data-testid="stFileUploader"] section button {
    background: rgba(255, 255, 255, 0.06) !important;
    color: #fafafa !important;
    border: 1px solid rgba(255, 255, 255, 0.12) !important;
    border-radius: 6px !important;
    box-shadow: none !important;
}

/* Input fields */
input, textarea, [data-baseweb="input"], [data-baseweb="select"] {
    background-color: #141414 !important;
    border: 1px solid rgba(255, 255, 255, 0.12) !important;
    color: #fafafa !important;
    border-radius: 8px !important;
    box-shadow: none !important;
}

/* Tabs */
[data-baseweb="tab-list"] {
    background: transparent !important;
    border-bottom: 1px solid rgba(255, 255, 255, 0.10) !important;
}

[data-baseweb="tab"] {
    color: #a3a3a3 !important;
    font-family: 'Manrope', sans-serif !important;
}

[data-baseweb="tab"][aria-selected="true"] {
    color: #fafafa !important;
    border-bottom: 2px solid #ffedd5 !important;
}

/* Dataframe container */
[data-testid="stDataFrame"] {
    border: 1px solid rgba(255, 255, 255, 0.10) !important;
    border-radius: 8px !important;
    background: #0e0e0e !important;
    box-shadow: none !important;
}

/* Expanders */
details, [data-testid="stExpander"] {
    border: 1px solid rgba(255, 255, 255, 0.10) !important;
    border-radius: 8px !important;
    background: rgba(255, 255, 255, 0.02) !important;
    box-shadow: none !important;
}

hr {
    border-color: rgba(255, 255, 255, 0.10) !important;
}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# SECTION 3: MODELS AND CONFIGURATION
# ==============================================================================

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
SHEET_STORE         = "Store"
SHEET_SKU           = "SKU"

OUTPUT_SHEET_PLANOGRAM  = "Generated Planogram"
OUTPUT_SHEET_VALIDATION = "Validation"
OUTPUT_SHEET_AUDIT      = "Audit Report"
OUTPUT_SHEET_STORE_SKU  = "Store SKU"

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

# ==============================================================================
# SECTION 4: LOGGER AND PARSER
# ==============================================================================

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


_RE_LEADING_COUNT = re.compile(r"^\s*(\d+)\s+bay\s*[-–]?\s*", re.IGNORECASE)
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

# ==============================================================================
# SECTION 5: WORKBOOK LOADER AND ALLOCATOR
# ==============================================================================

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
        logger.warning(f"Invalid Facing '{raw}' for SKU {sku}: defaulting to 1.", store=store)
        return 1
    if val not in VALID_FACINGS:
        logger.warning(f"Facing '{raw}' for SKU {sku} not in {{1,2}}: defaulting to 1.", store=store)
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
            logger.warning(f"Unsupported Bay Size '{bay_size}' in Bay {bay_num}: skipping.", store=store, bay_num=bay_num)
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

# ==============================================================================
# SECTION 6: AUDIT ENGINE
# ==============================================================================

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
            'Passed': bool(store_pass),
            'Details': " | ".join(cat_statuses)
        })

    return pd.DataFrame(results)

# ==============================================================================
# SECTION 7: STORE SKU GENERATOR
# ==============================================================================

_SKU_SHEET_NON_BOARDSET_COLS = {
    "sku", "description", "rank", "cf id", "cf", "color flow",
    "store", "total", "facings", "facing", "status",
}


def generate_store_sku_sheet(file_bytes: bytes) -> bytes:
    try:
        xl = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    except Exception as exc:
        raise RuntimeError(f"Cannot open workbook: {exc}") from exc

    def _find_sheet(candidates: List[str]) -> str:
        for name in xl.sheet_names:
            if name.strip().lower() in [c.lower() for c in candidates]:
                return name
        raise ValueError(f"Sheet matching {candidates} not found. Available sheets: {xl.sheet_names}")

    store_sheet_name = _find_sheet(["Store", "Stores"])
    df_store = xl.parse(store_sheet_name, dtype=str)
    df_store.columns = [str(c).strip() for c in df_store.columns]
    df_store.dropna(how="all", inplace=True)
    df_store.reset_index(drop=True, inplace=True)

    if df_store.empty:
        raise ValueError("The 'Store' sheet is empty.")

    store_cols_list = list(df_store.columns)
    store_num_col = store_cols_list[0]
    board_set_col = store_cols_list[1] if len(store_cols_list) > 1 else None

    if board_set_col is None:
        raise ValueError("The 'Store' sheet must have at least 2 columns: Store and SO Board Set.")

    board_set_to_stores: Dict[str, List[str]] = defaultdict(list)
    for _, row in df_store.iterrows():
        store_num = _clean_store_id(row[store_num_col])
        board_set = _str(row[board_set_col])
        if store_num and board_set:
            board_set_to_stores[board_set.strip()].append(store_num)

    if not board_set_to_stores:
        raise ValueError("No valid Store to SO Board Set mappings found in the 'Store' sheet.")

    sku_sheet_name = _find_sheet(["SKU", "SKUs"])
    df_sku = xl.parse(sku_sheet_name, dtype=str)
    df_sku.columns = [str(c).strip() for c in df_sku.columns]
    df_sku.dropna(how="all", inplace=True)
    df_sku.reset_index(drop=True, inplace=True)

    if df_sku.empty:
        raise ValueError("The 'SKU' sheet is empty.")

    sku_cols = list(df_sku.columns)
    sku_col = sku_cols[0]
    desc_col = sku_cols[1] if len(sku_cols) > 1 else None

    boardset_cols: List[str] = [
        col for col in sku_cols[2:]
        if col.strip().lower() not in _SKU_SHEET_NON_BOARDSET_COLS
    ]

    if not boardset_cols:
        raise ValueError("No SO Board Set columns detected in the 'SKU' sheet.")

    output_rows: List[Dict[str, Any]] = []

    for _, row in df_sku.iterrows():
        sku = _str(row[sku_col])
        desc = _str(row[desc_col]) if desc_col else ""
        if not sku:
            continue

        for bs_col in boardset_cols:
            cell_val = _str(row.get(bs_col, ""))
            if not cell_val.strip().upper().startswith("Y"):
                continue

            stores_for_bs = board_set_to_stores.get(bs_col.strip(), [])
            for store in stores_for_bs:
                output_rows.append({
                    "Store": store,
                    "SKU": sku,
                    "Description": desc,
                    "Store-SKU": f"{sku}-{store}",
                    "Facings": 1,
                    "SKU - Facing": f"{sku}-1",
                })

    if not output_rows:
        raise ValueError("No output rows generated. Check that SKU sheet contains Y-marked cells matching SO Board Sets.")

    def _sort_key(r: Dict[str, Any]):
        try:
            return (int(r["Store"]), r["SKU"])
        except (ValueError, TypeError):
            return (0, r["SKU"])

    output_rows.sort(key=_sort_key)
    out_df = pd.DataFrame(output_rows, columns=["Store", "SKU", "Description", "Store-SKU", "Facings", "SKU - Facing"])

    wb = _openpyxl_load(BytesIO(file_bytes))
    if OUTPUT_SHEET_STORE_SKU in wb.sheetnames:
        del wb[OUTPUT_SHEET_STORE_SKU]

    ws_out = wb.create_sheet(title=OUTPUT_SHEET_STORE_SKU)
    ws_out.append(list(out_df.columns))
    for _, row in out_df.iterrows():
        ws_out.append([row[col] for col in out_df.columns])

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(fill_type="solid", fgColor="171717")
    header_align = Alignment(horizontal="center", vertical="center")
    for cell in ws_out[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
    ws_out.row_dimensions[1].height = 22

    for col_cells in ws_out.columns:
        length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col_cells)
        ws_out.column_dimensions[get_column_letter(col_cells[0].column)].width = min(length + 3, 50)
    ws_out.freeze_panes = "A2"

    output_bio = BytesIO()
    wb.save(output_bio)
    output_bio.seek(0)
    return output_bio.read()

# ==============================================================================
# SECTION 8: EXCEL WRITER (WITH AUDIT REPORT SHEET)
# ==============================================================================

_HEADER_FONT = Font(bold=True, color="FFFFFF")
_DARK_HEADER = PatternFill(fill_type="solid", fgColor="171717")
_ERROR_HEADER = PatternFill(fill_type="solid", fgColor="450A0A")
_AUDIT_HEADER = PatternFill(fill_type="solid", fgColor="1E293B")
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _style_header(ws, fill: PatternFill = _DARK_HEADER) -> None:
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
    _style_header(ws_pog, fill=_DARK_HEADER)
    _auto_width(ws_pog)
    ws_pog.freeze_panes = "A2"

    val_cols: List[str] = ["Level", "Store", "Bay#", "Message"]
    ws_val = wb.create_sheet(title=OUTPUT_SHEET_VALIDATION)
    ws_val.append(val_cols)
    for _, row in validation_df.iterrows():
        ws_val.append([_str(row.get(c, "")) for c in val_cols])
    _style_header(ws_val, fill=_ERROR_HEADER)
    _auto_width(ws_val)
    ws_val.freeze_panes = "A2"

    if audit_df is not None and not audit_df.empty:
        ws_aud = wb.create_sheet(title=OUTPUT_SHEET_AUDIT)
        aud_cols = ["Store", "Status", "Details"]
        ws_aud.append(aud_cols)
        for _, row in audit_df.iterrows():
            ws_aud.append([_str(row.get(c, "")) for c in aud_cols])
        _style_header(ws_aud, fill=_AUDIT_HEADER)
        _auto_width(ws_aud)
        ws_aud.freeze_panes = "A2"

    output_bio = BytesIO()
    wb.save(output_bio)
    output_bio.seek(0)
    return output_bio.read()

# ==============================================================================
# SECTION 9: STREAMLIT UI WORKFLOWS
# ==============================================================================

st.markdown("""
<div class="chai-hero">
  <div class="chai-hero-title">HD Planogram Generator</div>
  <p class="chai-hero-sub">Generate bay layouts and verify SKU positions with <span class="highlight">automated compliance</span> audit validation.</p>
</div>
""", unsafe_allow_html=True)

st.sidebar.markdown('<div class="chai-section-hdr">Workflow selection</div>', unsafe_allow_html=True)
app_mode = st.sidebar.radio(
    "Choose workflow",
    ["Generate planogram and run audit", "Audit existing planogram", "Generate store SKU sheet"],
    index=0
)

# Workflow 1: Generate planogram and run audit
if app_mode == "Generate planogram and run audit":
    st.sidebar.markdown("---")
    st.sidebar.markdown('<div class="chai-section-hdr">Audit rule settings</div>', unsafe_allow_html=True)
    notes_filter = st.sidebar.text_input("Store notes filter (regex)", "baja|vigo")
    keywords_str = st.sidebar.text_input("Target SKU keywords", "BAJA, VIGO")
    keywords = [k.strip() for k in keywords_str.split(",") if k.strip()]
    categories = ["Display", "SO", "Stock"]

    st.markdown("""
    <div class="chai-card">
      <strong>Planogram generation and audit workflow:</strong><br>
      Upload your raw workbook containing <code>Store List</code>, <code>Stock SKUs and Displays</code>, and <code>Special Order Boards</code>.
      The engine executes <span class="highlight">precision bay allocation</span>, validates facing rules, and audits target SKU end placement.
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Upload raw workbook (.xlsx)", type=["xlsx", "xls"])
    st.markdown("")
    run_clicked = st.button("Generate planogram and run audit ↗", type="primary")

    if run_clicked:
        if uploaded_file is None:
            st.error("Please upload an Excel workbook first.")
            st.stop()

        file_bytes = uploaded_file.getvalue()
        logger     = PlanogramLogger()

        try:
            with st.spinner("Loading workbook and parsing sheets..."):
                wb_data = load_workbook_from_bytes(file_bytes, logger)

            with st.spinner("Allocating SKUs and generating bay layout..."):
                planogram_df, validation_df = generate_planogram(wb_data, logger)

            with st.spinner("Running automated audit verification..."):
                audit_df = audit_planogram_df(wb_data.store_list, planogram_df, notes_filter, keywords, categories)

            with st.spinner("Bundling planogram and audit report..."):
                output_bytes = write_output_bytes(file_bytes, planogram_df, validation_df, audit_df)

        except Exception as exc:
            st.error(f"Processing error: {exc}")
            with st.expander("Technical error details"):
                st.code(traceback.format_exc())
            st.stop()

        st.markdown("---")
        st.markdown('<div class="chai-section-hdr">Generation and audit complete</div>', unsafe_allow_html=True)

        n_stores = planogram_df["Store"].nunique() if not planogram_df.empty else 0
        n_rows   = len(planogram_df)
        n_issues = len(validation_df)

        filtered_count = len(audit_df)
        pass_count = int(audit_df['Passed'].sum()) if len(audit_df) > 0 else 0
        pass_rate = (pass_count / filtered_count * 100) if filtered_count > 0 else 0.0

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f'<div class="chai-stat-card"><div class="chai-stat-value">{n_stores:,}</div><div class="chai-stat-label">Stores processed</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="chai-stat-card"><div class="chai-stat-value">{n_rows:,}</div><div class="chai-stat-label">Planogram rows</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="chai-stat-card"><div class="chai-stat-value highlight-green">{pass_count}/{filtered_count}</div><div class="chai-stat-label">Audited stores passed</div></div>', unsafe_allow_html=True)
        with c4:
            st.markdown(f'<div class="chai-stat-card"><div class="chai-stat-value highlight-green">{pass_rate:.1f}%</div><div class="chai-stat-label">Compliance rate</div></div>', unsafe_allow_html=True)

        st.markdown("")
        st.download_button(
            label="Download audited planogram (.xlsx)",
            data=output_bytes,
            file_name="Planogram_Output_Audited.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("---")
        tab1, tab2, tab3 = st.tabs(["Store audit report", "Planogram preview", "Validation issues"])

        with tab1:
            st.markdown('<div class="chai-section-hdr">Audit results breakdown</div>', unsafe_allow_html=True)
            st.dataframe(audit_df[['Store', 'Status', 'Details']], use_container_width=True, hide_index=True)

        with tab2:
            st.markdown('<div class="chai-section-hdr">Generated planogram preview</div>', unsafe_allow_html=True)
            st.dataframe(planogram_df.head(200), use_container_width=True, hide_index=True)

        with tab3:
            st.markdown('<div class="chai-section-hdr">Validation log</div>', unsafe_allow_html=True)
            if n_issues > 0:
                st.dataframe(validation_df, use_container_width=True, hide_index=True)
            else:
                st.success("No validation issues found.")

# Workflow 2: Audit existing planogram only
elif app_mode == "Audit existing planogram":
    st.sidebar.markdown("---")
    st.sidebar.markdown('<div class="chai-section-hdr">Audit rule settings</div>', unsafe_allow_html=True)
    notes_filter = st.sidebar.text_input("Store notes filter (regex)", "baja|vigo")
    keywords_str = st.sidebar.text_input("Target SKU keywords", "BAJA, VIGO")
    keywords = [k.strip() for k in keywords_str.split(",") if k.strip()]
    categories = ["Display", "SO", "Stock"]

    st.markdown("""
    <div class="chai-card">
      <strong>Standalone audit workflow:</strong><br>
      Upload an existing planogram workbook containing <code>Store List</code> and <code>Generated Planogram</code> sheets to audit <span class="highlight">keyword SKU placement</span> across all forced stores.
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Upload existing planogram workbook (.xlsx)", type=["xlsx", "xls"])
    if uploaded_file is not None:
        xl = pd.ExcelFile(uploaded_file)
        store_sheet = st.selectbox("Store metadata sheet", xl.sheet_names, index=0 if "Store List" in xl.sheet_names else 0)
        pog_sheet = st.selectbox("Generated planogram sheet", xl.sheet_names, index=xl.sheet_names.index("Generated Planogram") if "Generated Planogram" in xl.sheet_names else min(1, len(xl.sheet_names)-1))

        st.markdown("")
        if st.button("Run compliance audit ↗", type="primary"):
            df_s = pd.read_excel(uploaded_file, sheet_name=store_sheet)
            df_p = pd.read_excel(uploaded_file, sheet_name=pog_sheet)

            audit_df = audit_planogram_df(df_s, df_p, notes_filter, keywords, categories)

            st.markdown("---")
            st.markdown('<div class="chai-section-hdr">Audit results</div>', unsafe_allow_html=True)

            filtered_count = len(audit_df)
            pass_count = int(audit_df['Passed'].sum()) if len(audit_df) > 0 else 0
            pass_rate = (pass_count / filtered_count * 100) if filtered_count > 0 else 0.0

            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f'<div class="chai-stat-card"><div class="chai-stat-value">{filtered_count}</div><div class="chai-stat-label">Stores filtered</div></div>', unsafe_allow_html=True)
            with c2:
                st.markdown(f'<div class="chai-stat-card"><div class="chai-stat-value highlight-green">{pass_count}</div><div class="chai-stat-label">Passed</div></div>', unsafe_allow_html=True)
            with c3:
                st.markdown(f'<div class="chai-stat-card"><div class="chai-stat-value highlight-green">{pass_rate:.1f}%</div><div class="chai-stat-label">Compliance rate</div></div>', unsafe_allow_html=True)

            st.markdown("")
            st.dataframe(audit_df[['Store', 'Status', 'Details']], use_container_width=True, hide_index=True)

            csv_data = audit_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="Download audit results (.csv)",
                data=csv_data,
                file_name="Planogram_Audit_Results.csv",
                mime="text/csv",
            )

# Workflow 3: Generate store SKU sheet
else:
    st.markdown("""
    <div class="chai-card">
      <strong>Store SKU generator workflow:</strong><br>
      Upload a workbook containing <code>Store</code> and <code>SKU</code> sheets to generate a unified <span class="highlight">Store SKU mapping</span> sheet for all active board set stores.
    </div>
    """, unsafe_allow_html=True)

    sku_uploaded_file = st.file_uploader("Upload store SKU workbook (.xlsx)", type=["xlsx", "xls"], key="sku_uploader")
    st.markdown("")
    sku_run_clicked = st.button("Generate store SKU sheet ↗", type="primary", key="sku_btn")

    if sku_run_clicked:
        if sku_uploaded_file is None:
            st.error("Please upload an Excel workbook first.")
            st.stop()

        sku_file_bytes = sku_uploaded_file.getvalue()

        try:
            with st.spinner("Processing store SKU mappings..."):
                sku_output_bytes = generate_store_sku_sheet(sku_file_bytes)

        except Exception as exc:
            st.error(f"Processing error: {exc}")
            with st.expander("Technical details"):
                st.code(traceback.format_exc())
            st.stop()

        st.markdown("---")
        st.markdown('<div class="chai-section-hdr">Generation complete</div>', unsafe_allow_html=True)
        st.success("Store SKU sheet generated successfully.")

        st.markdown("")
        st.download_button(
            label="Download store SKU output (.xlsx)",
            data=sku_output_bytes,
            file_name="Store_SKU_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="sku_dl_btn",
        )

st.markdown("---")
st.markdown('<p style="text-align:center;color:#737373;font-size:0.85rem;margin-top:1.5rem;">HD Planogram Generator</p>', unsafe_allow_html=True)
