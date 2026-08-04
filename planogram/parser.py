from __future__ import annotations
import re
from typing import Any, List, Tuple

from planogram.logger import PlanogramLogger

# ---------------------------------------------------------------------------
# Pre-compiled regexes (compiled once at import time for performance)
# ---------------------------------------------------------------------------

# Matches the leading "N Bay -", "N BAY -", "N Bay " prefix.
#
# The dash after "Bay" is OPTIONAL (? quantifier) to handle real-world
# POG strings that omit it, e.g.:
#   "8 Bay 99,99,..."                  → no dash between Bay and tokens
#   "1 Bay 99C - LFT - Full(...)"      → no dash, single token with description
#   "1 Bay 99 - OPP Plug & Play"       → no dash before token
#   "6 Bay-99,99,..." (dash attached)  → still handled by the dash clause
#   "8 Bay - 99 - description"         → standard format with dash + spaces
_RE_LEADING_COUNT = re.compile(
    r"^\s*(\d+)\s+bay\s*[-\u2013]?\s*",
    re.IGNORECASE,
)

# Matches a COMPLETE clean token with no trailing description text:
#   "99"   "99C"   "87"   "51"       (plain bay sizes)
#   "8-99" "2-51"                     (multiplier pair, FORMAT A)
_RE_VALID_TOKEN = re.compile(r"^[A-Za-z0-9]+(?:-[A-Za-z0-9]+)?$")

# Matches a multiplier pair: "<count>-<size>"  e.g. "8-99", "2-51"
_RE_MULTIPLIER_TOKEN = re.compile(r"^(\d+)-([A-Za-z0-9]+)$")

# Extracts ONLY the leading numeric bay-size from a token that has trailing
# description text.  Restricted to: digits + up to 2 letters.
#
#   "99 - Floor Tile Test - Str 1602"  →  "99"
#   "99C - Stone Look..."              →  "99C"
#   "87 - Mid..."                      →  "87"
#   "99-Stone Look -Mid -v1 - Apr"     →  "99"  (stops before the dash)
#
# Previously used r"^([A-Za-z0-9]+(?:-[A-Za-z0-9]+)?)" which incorrectly
# captured "99-STONE" as a single bay-size token (Bug #3).
_RE_BAYSIZE_PREFIX = re.compile(r"^(\d+[A-Za-z]{0,2})")


def parse_pog_string(
    pog: str,
    store: Any = "",
    logger: PlanogramLogger | None = None,
) -> List[str]:
    """Parse a POG/LFT description string into an ordered list of bay-size tokens.

    Three formats are auto-detected:
        FORMAT A — multiplier shorthand  : "10 Bay - 8-99,2-51"
        FORMAT B — explicit comma-list   : "7 BAY - 99,99,99,99,99,99,87"
        FORMAT C — single repeated size  : "8 Bay - 99"
                                           "12 Bay - 99 - Floor Tile Test - Str 1602"

    Trailing description text after the bay tokens is ALWAYS silently discarded
    regardless of format.

    Parameters
    ----------
    pog:
        Raw POG or LFT string from the workbook.
    store:
        Store ID used only in log messages.
    logger:
        PlanogramLogger instance, or None to suppress warnings.

    Returns
    -------
    List[str]
        Ordered bay-size tokens, e.g. ["99"] * 12.
        Empty list if the string cannot be parsed.
    """
    if not isinstance(pog, str) or not pog.strip():
        return []

    pog = pog.strip()

    # ── Step 1: Strip the leading "N Bay -" prefix ─────────────────────────
    m_count = _RE_LEADING_COUNT.match(pog)
    if not m_count:
        if logger:
            logger.warning(f"Cannot parse bay count from POG: '{pog}'", store=store)
        return []

    declared_count: int = int(m_count.group(1))
    remainder: str = pog[m_count.end():].strip()

    # ── Step 2: Collect bay-size tokens from comma-separated pieces ────────
    #
    # Three cases per comma-split token:
    #   (a) Clean token ("99", "99C", "8-99") → matches _RE_VALID_TOKEN → append, continue
    #   (b) Token with trailing description   → extract prefix via _RE_BAYSIZE_PREFIX → append, stop
    #   (c) Unrecognisable token              → stop immediately
    #
    # Stopping on case (b)/(c) is correct because description text always
    # signals the end of the bay-size list.

    raw_tokens = [t.strip() for t in remainder.split(",")]
    bay_tokens_raw: List[str] = []

    for tok in raw_tokens:
        if _RE_VALID_TOKEN.match(tok):
            # Case (a): pure bay-size or multiplier token — keep and continue
            bay_tokens_raw.append(tok)
        else:
            # Case (b): token has trailing description; extract the prefix only
            m_prefix = _RE_BAYSIZE_PREFIX.match(tok)
            if m_prefix:
                bay_tokens_raw.append(m_prefix.group(1))
            # Always stop — description text means no more bay tokens follow
            break

    if not bay_tokens_raw:
        if logger:
            logger.warning(f"No bay tokens found in POG: '{pog}'", store=store)
        return []

    # ── Step 3: Detect format ──────────────────────────────────────────────
    all_multiplier = all(_RE_MULTIPLIER_TOKEN.match(tok) for tok in bay_tokens_raw)

    if all_multiplier:
        # FORMAT A: "8-99,2-51"  →  ["99"]*8 + ["51"]*2
        expanded: List[str] = []
        for tok in bay_tokens_raw:
            m = _RE_MULTIPLIER_TOKEN.match(tok)
            count_str, size_str = m.group(1), m.group(2)
            expanded.extend([size_str.upper()] * int(count_str))
    elif len(bay_tokens_raw) == 1:
        # FORMAT C: single size token → repeat declared_count times.
        # Handles both plain "8 Bay - 99" AND
        # "12 Bay - 99 - Floor Tile Test - Str 1602 - June 26".
        expanded = [bay_tokens_raw[0].upper()] * declared_count
    else:
        # FORMAT B: explicit comma-list → use as-is
        expanded = [tok.upper() for tok in bay_tokens_raw]

    # ── Step 4: Warn if parsed count ≠ declared count ─────────────────────
    if len(expanded) != declared_count and logger:
        logger.warning(
            f"Parsed {len(expanded)} bay(s) but header declares {declared_count}: '{pog}'",
            store=store,
        )

    return expanded


# ---------------------------------------------------------------------------
# Self-tests — call run_self_tests() to verify the parser is working
# ---------------------------------------------------------------------------
_SELF_TESTS: List[Tuple[str, List[str]]] = [
    # FORMAT C — standard with dash
    ("8 Bay - 99",
     ["99"] * 8),
    # FORMAT C — WITH trailing description text
    ("12 Bay - 99 - Floor Tile Test - Str 1602 - June 26",
     ["99"] * 12),
    ("5 Bay - 99 - Stone Look Tile -South-v1 - April 2026",
     ["99"] * 5),
    ("1 Bay - 51 - LFT - Full(Sku Swap)",
     ["51"]),
    # FORMAT B — clean explicit list with dash
    ("7 BAY - 99,99,99,99,99,99,87",
     ["99"] * 6 + ["87"]),
    # FORMAT B — explicit list WITH trailing description on last token
    ("10 Bay - 99,99,99,99,99,99,99,99,99,87 - Floor Tile Test - Str 1503 - June 26",
     ["99"] * 9 + ["87"]),
    ("7 BAY - 99,99,99,99,99,99,87 - Stone Look Tile - Mid",
     ["99"] * 6 + ["87"]),
    # FORMAT B — with bay-size variants
    ("10 BAY - 99C,99,99,99,99,99C,99,99,99,99",
     ["99C", "99", "99", "99", "99", "99C", "99", "99", "99", "99"]),
    # FORMAT A — multiplier shorthand
    ("10 Bay - 8-99,2-51",
     ["99"] * 8 + ["51"] * 2),
    # ── NEW: Real-world failure cases (Bug fixes) ─────────────────────────
    # Bug 1: NO dash between 'Bay' and the token list; trailing description
    ("8 Bay 99,99,99,99,99,99,87,87 -Stone Look Tile- Mid - April 2026",
     ["99"] * 6 + ["87"] * 2),
    # Bug 2: NO dash between 'Bay' and single token; LFT-style description
    ("1 Bay 99C - LFT - Full(Shoregaze Sku Swap) - April 2026",
     ["99C"]),
    # Bug 3: Dash attached to 'Bay' (no space); last token has 'word-word' description
    #         _RE_BAYSIZE_PREFIX must NOT capture '99-STONE'; only '99'
    ("6 Bay-99,99,99,75,99,99-Stone Look -Mid -v1 - April 2026",
     ["99", "99", "99", "75", "99", "99"]),
    # Bug 5: NO dash between 'Bay' and token; description contains '&' and spaces
    ("1 Bay 99 - OPP Plug & Play",
     ["99"]),
    # Edge: dash attached to Bay, no space, single-size format
    ("5 Bay-99 - Stone Look Tile - April 2026",
     ["99"] * 5),
]


def run_self_tests() -> bool:
    """Run all parser self-tests. Returns True if every test passes."""
    all_pass = True
    for pog_str, expected in _SELF_TESTS:
        result = parse_pog_string(pog_str, store="SELF-TEST")
        if result != expected:
            print(f"FAIL: '{pog_str}'")
            print(f"  Expected: {expected}")
            print(f"  Got:      {result}")
            all_pass = False
    if all_pass:
        print(f"POG parser: all {len(_SELF_TESTS)} self-tests passed.")
    return all_pass
