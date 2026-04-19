"""
file_merger.py
Multi-file comparison and merge logic.

Workflow:
  - Master file : full history report — all action items (open + closed)
  - New export  : latest snapshot — only currently open items

merge_files() produces an updated master where:
  1. Rows in master but NOT in new export  → Status column set to closed_value
  2. Rows in new export but NOT in master  → appended as new open rows
  3. Rows present in both                 → left unchanged

Matching is done on a single unique ID column chosen by the user.
"""
from __future__ import annotations

import pandas as pd


# ---------------------------------------------------------------------------
# Public types
# ---------------------------------------------------------------------------

class MergeConfig:
    """
    Parameters
    ----------
    id_col        : column whose value uniquely identifies a row in both files
    status_col    : column in the master that holds Open/Closed status
    closed_value  : string to write when marking a row Closed  (default "Closed")
    open_value    : string that represents an Open item         (default "Open")
    """
    def __init__(
        self,
        id_col: str,
        status_col: str,
        closed_value: str = "Closed",
        open_value: str = "Open",
    ) -> None:
        self.id_col = id_col
        self.status_col = status_col
        self.closed_value = closed_value
        self.open_value = open_value


class MergeResult:
    """
    Attributes
    ----------
    df          : updated master DataFrame (input not mutated)
    closed_ids  : IDs that were newly marked Closed
    added_ids   : IDs that were appended as new rows
    """
    def __init__(
        self,
        df: pd.DataFrame,
        closed_ids: list,
        added_ids: list,
    ) -> None:
        self.df = df
        self.closed_ids = closed_ids
        self.added_ids = added_ids

    @property
    def summary(self) -> str:
        parts = []
        if self.closed_ids:
            parts.append(f"🔒 **{len(self.closed_ids)}** item(s) marked Closed")
        if self.added_ids:
            parts.append(f"➕ **{len(self.added_ids)}** new item(s) added")
        if not parts:
            parts.append("✅ No changes — master is already up to date")
        return "\n\n".join(parts)


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def validate_columns(
    master: pd.DataFrame,
    new_export: pd.DataFrame,
    cfg: MergeConfig,
) -> list[str]:
    """Return a list of human-readable error strings; empty = valid."""
    errors: list[str] = []
    if cfg.id_col not in master.columns:
        errors.append(f"ID column '{cfg.id_col}' not found in master file.")
    if cfg.id_col not in new_export.columns:
        errors.append(f"ID column '{cfg.id_col}' not found in new export file.")
    if cfg.status_col not in master.columns:
        errors.append(
            f"Status column '{cfg.status_col}' not found in master file — "
            "it will be created automatically, but verify this is intentional."
        )
    return errors


# ---------------------------------------------------------------------------
# Core merge
# ---------------------------------------------------------------------------

def merge_files(
    master: pd.DataFrame,
    new_export: pd.DataFrame,
    cfg: MergeConfig,
) -> MergeResult:
    """
    Perform the merge. Does NOT mutate either input DataFrame.
    """
    result = master.copy()

    # Ensure status column exists
    if cfg.status_col not in result.columns:
        result[cfg.status_col] = cfg.open_value

    # Build ID sets (string-normalised for safe comparison)
    master_ids: dict[str, int] = {
        str(result.at[idx, cfg.id_col]).strip(): idx
        for idx in result.index
    }
    new_ids: set[str] = {
        str(row[cfg.id_col]).strip()
        for _, row in new_export.iterrows()
    }

    closed_ids: list = []
    added_ids: list = []

    # ── 1. Mark rows absent from new export as Closed ────────────────────────
    for id_val, idx in master_ids.items():
        if id_val not in new_ids:
            current = str(result.at[idx, cfg.status_col]).strip()
            if current != cfg.closed_value:          # skip already-closed rows
                result.at[idx, cfg.status_col] = cfg.closed_value
                closed_ids.append(id_val)

    # ── 2. Append rows that are new in the export ─────────────────────────────
    new_rows: list[dict] = []
    for _, row in new_export.iterrows():
        id_val = str(row[cfg.id_col]).strip()
        if id_val not in master_ids:
            new_row: dict = {}
            for col in result.columns:
                if col in new_export.columns:
                    new_row[col] = row[col]
                elif col == cfg.status_col:
                    new_row[col] = cfg.open_value
                else:
                    new_row[col] = pd.NA
            new_rows.append(new_row)
            added_ids.append(id_val)

    if new_rows:
        result = pd.concat(
            [result, pd.DataFrame(new_rows, columns=result.columns)],
            ignore_index=True,
        )

    return MergeResult(df=result, closed_ids=closed_ids, added_ids=added_ids)
