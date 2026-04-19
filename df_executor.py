"""
df_executor.py
Executes LLM-generated Python commands against a Pandas DataFrame.

Security measures:
  1. Every code string is AST-checked by sandbox.check_code() BEFORE exec().
  2. exec() runs with __builtins__ explicitly set to a minimal safe dict,
     preventing access to dangerous built-ins even if the AST check is bypassed.
  3. A row-count / memory guard refuses DataFrames that grow unreasonably.
"""
import numpy as np
import pandas as pd

from sandbox import check_code

# ---------------------------------------------------------------------------
# Limits
# ---------------------------------------------------------------------------
MAX_ROWS = 500_000          # refuse result DataFrames larger than this
MAX_MEMORY_MB = 512         # refuse result DataFrames heavier than this

# ---------------------------------------------------------------------------
# Minimal safe __builtins__ exposed to exec()
# Only pure-value built-ins that DataFrame transformations legitimately need.
# ---------------------------------------------------------------------------
_SAFE_BUILTINS: dict = {
    # types / constructors
    "int": int, "float": float, "str": str, "bool": bool,
    "list": list, "dict": dict, "tuple": tuple, "set": set,
    "frozenset": frozenset, "bytes": bytes,
    # numeric helpers
    "abs": abs, "round": round, "min": min, "max": max,
    "sum": sum, "len": len, "range": range, "enumerate": enumerate,
    "zip": zip, "map": map, "filter": filter, "sorted": sorted,
    "reversed": reversed, "any": any, "all": all,
    # safe inspection
    "isinstance": isinstance, "issubclass": issubclass,
    "repr": repr, "print": print,
    # exceptions (needed for try/except in generated code)
    "Exception": Exception, "ValueError": ValueError,
    "TypeError": TypeError, "KeyError": KeyError,
    "IndexError": IndexError, "AttributeError": AttributeError,
    # constants
    "True": True, "False": False, "None": None,
}

# Globals injected into every exec() call — no __builtins__ leakage
_EXEC_GLOBALS: dict = {
    "__builtins__": _SAFE_BUILTINS,
    "np": np,
    "pd": pd,
}


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _patch_deprecated(cmd: str) -> tuple[str, str | None]:
    """Replace deprecated df.append(); returns (patched_cmd, warning | None)."""
    if "df.append" in cmd:
        patched = (
            cmd.replace("df.append", "pd.concat([df,")
               .replace("]", "], ignore_index=True)")
        )
        return patched, "Replaced deprecated df.append() with pd.concat()."
    return cmd, None


def _guard_result(df: pd.DataFrame) -> None:
    """Raise RuntimeError if the resulting DataFrame is unreasonably large."""
    if len(df) > MAX_ROWS:
        raise RuntimeError(
            f"Result DataFrame has {len(df):,} rows, which exceeds the "
            f"safety limit of {MAX_ROWS:,}. Transformation blocked."
        )
    mem_mb = df.memory_usage(deep=True).sum() / (1024 ** 2)
    if mem_mb > MAX_MEMORY_MB:
        raise RuntimeError(
            f"Result DataFrame uses {mem_mb:.1f} MB, which exceeds the "
            f"safety limit of {MAX_MEMORY_MB} MB. Transformation blocked."
        )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def run_commands(
    df: pd.DataFrame, commands: list[str]
) -> tuple[pd.DataFrame, list[str]]:
    """
    Security-check and execute a list of single-line commands against a copy
    of `df`.

    Returns:
        (result_df, warnings)

    Raises:
        ValueError   – if any command fails the security check.
        RuntimeError – if execution fails or the result exceeds size limits.
    """
    # ── 1. Security-check ALL commands before executing any ──────────────────
    full_code = "\n".join(commands)
    safe, reason = check_code(full_code)
    if not safe:
        raise ValueError(f"Code blocked by security check: {reason}")

    # ── 2. Execute inside restricted sandbox ─────────────────────────────────
    local_vars: dict = {"df": df.copy()}
    warnings: list[str] = []

    for cmd in commands:
        cmd, warning = _patch_deprecated(cmd)
        if warning:
            warnings.append(warning)
        try:
            exec(cmd, _EXEC_GLOBALS, local_vars)  # noqa: S102
        except Exception as exc:
            raise RuntimeError(f"Execution error in command:\n  {cmd}\n{exc}") from exc

    # ── 3. Guard result size ─────────────────────────────────────────────────
    result: pd.DataFrame = local_vars["df"]
    _guard_result(result)

    return result, warnings


def run_single(
    df: pd.DataFrame, code: str
) -> tuple[pd.DataFrame, list[str]]:
    """Convenience wrapper: splits a multi-line code block into commands."""
    commands = [line for line in code.splitlines() if line.strip()]
    return run_commands(df, commands)
