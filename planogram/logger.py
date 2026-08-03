from __future__ import annotations
from typing import Any, Dict, List


class PlanogramLogger:
    """Collects validation issues during planogram generation.

    Each generation run should create a fresh instance so there is no
    shared state between uploads in the Streamlit app.
    """

    def __init__(self) -> None:
        self.issues: List[Dict[str, Any]] = []
        self._seen: set = set()

    def _append(
        self, level: str, msg: str, store: Any = "", bay_num: Any = ""
    ) -> None:
        # Deduplicate identical messages for the same store/bay so we don't
        # flood the output when the same problem repeats across many rows
        # (e.g. invalid Facing value for thousands of SKUs).
        dedup_key = (level, str(store), str(bay_num), msg)
        if dedup_key in self._seen:
            return
        self._seen.add(dedup_key)
        self.issues.append(
            {
                "Level":   level,
                "Store":   str(store),
                "Bay#":    str(bay_num),
                "Message": msg,
            }
        )

    def warning(self, msg: str, store: Any = "", bay_num: Any = "") -> None:
        self._append("WARNING", msg, store, bay_num)

    def error(self, msg: str, store: Any = "", bay_num: Any = "") -> None:
        self._append("ERROR", msg, store, bay_num)

    def info(self, msg: str) -> None:
        pass  # INFO messages are not surfaced in the Validation sheet
