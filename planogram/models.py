from __future__ import annotations
from dataclasses import dataclass


@dataclass(frozen=True)
class BayRule:
    """Immutable capacity specification for one bay type."""

    display: int  # Display Board positions in this bay
    so: int       # Special Order Board positions in this bay
    stock: int    # Stock positions in this bay

    @property
    def total(self) -> int:
        """Total number of positions in this bay."""
        return self.display + self.so + self.stock


@dataclass
class SKURecord:
    """Represents one planogram position for a single SKU."""

    sku: str
    description: str
    facing: int    # Always 1 or 2 — values outside this range are clamped to 1
    sku_type: str  # "Display" | "SO" | "Stock"
