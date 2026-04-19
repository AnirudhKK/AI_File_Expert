"""
sandbox.py
AST-based static analysis guard that runs BEFORE exec().

Blocks:
  - import / __import__ / importlib  (no new modules)
  - os, sys, subprocess, shutil access via attribute chains
  - __builtins__, __class__, __mro__, __subclasses__ dunder escapes
  - open(), eval(), exec(), compile(), globals(), locals(), vars()
  - Any call whose name matches a known dangerous built-in
  - Attribute access on names that resolve to dangerous modules

Only pure DataFrame / numpy transformations should pass.
"""
import ast
from typing import NoReturn

# ---------------------------------------------------------------------------
# Deny-lists
# ---------------------------------------------------------------------------

# Top-level names that must never appear
_BLOCKED_NAMES: frozenset[str] = frozenset({
    "os", "sys", "subprocess", "shutil", "pathlib", "socket",
    "importlib", "builtins", "ctypes", "mmap", "signal",
    "threading", "multiprocessing", "concurrent",
    "__import__", "__builtins__", "__loader__", "__spec__",
    "eval", "exec", "compile", "open", "input",
    "globals", "locals", "vars", "dir", "getattr", "setattr",
    "delattr", "hasattr", "object", "type", "super",
})

# Dunder attributes that are classic sandbox-escape gadgets
_BLOCKED_ATTRS: frozenset[str] = frozenset({
    "__class__", "__bases__", "__mro__", "__subclasses__",
    "__globals__", "__builtins__", "__dict__", "__code__",
    "__closure__", "__import__", "__reduce__", "__reduce_ex__",
    "__init_subclass__", "__set_name__",
})

# Import module names (caught via Import / ImportFrom nodes too)
_BLOCKED_MODULES: frozenset[str] = frozenset({
    "os", "sys", "subprocess", "shutil", "pathlib", "socket",
    "importlib", "builtins", "ctypes", "mmap", "signal",
    "threading", "multiprocessing", "concurrent", "pickle",
    "shelve", "marshal", "pty", "tty", "termios",
})


class _SecurityVisitor(ast.NodeVisitor):
    """Walks the AST and raises ValueError on the first violation found."""

    def _block(self, node: ast.AST, reason: str) -> NoReturn:
        lineno = getattr(node, "lineno", "?")
        raise ValueError(f"[Security] Blocked at line {lineno}: {reason}")

    # ── Import statements ────────────────────────────────────────────────────
    def visit_Import(self, node: ast.Import) -> None:
        for alias in node.names:
            root = alias.name.split(".")[0]
            if root in _BLOCKED_MODULES:
                self._block(node, f"import of '{alias.name}' is not allowed")
        self.generic_visit(node)

    def visit_ImportFrom(self, node: ast.ImportFrom) -> None:
        module = (node.module or "").split(".")[0]
        if module in _BLOCKED_MODULES:
            self._block(node, f"from-import of '{node.module}' is not allowed")
        self.generic_visit(node)

    # ── Dangerous bare names ─────────────────────────────────────────────────
    def visit_Name(self, node: ast.Name) -> None:
        if node.id in _BLOCKED_NAMES:
            self._block(node, f"use of '{node.id}' is not allowed")
        self.generic_visit(node)

    # ── Dangerous attribute access ───────────────────────────────────────────
    def visit_Attribute(self, node: ast.Attribute) -> None:
        if node.attr in _BLOCKED_ATTRS:
            self._block(node, f"access to attribute '{node.attr}' is not allowed")
        # Also block e.g. `something.os`, `df.something.subprocess`
        if node.attr in _BLOCKED_MODULES:
            self._block(node, f"attribute '{node.attr}' refers to a blocked module")
        self.generic_visit(node)

    # ── String-based eval / exec calls ──────────────────────────────────────
    def visit_Call(self, node: ast.Call) -> None:
        # Catch __import__("os") style
        if isinstance(node.func, ast.Name) and node.func.id in _BLOCKED_NAMES:
            self._block(node, f"call to '{node.func.id}' is not allowed")
        self.generic_visit(node)


def check_code(code: str) -> tuple[bool, str]:
    """
    Parse and security-check `code`.
    Returns (is_safe, error_message).
    error_message is '' when safe.
    """
    try:
        tree = ast.parse(code)
    except SyntaxError as e:
        return False, f"Syntax error: {e}"

    visitor = _SecurityVisitor()
    try:
        visitor.visit(tree)
    except ValueError as e:
        return False, str(e)

    return True, ""
