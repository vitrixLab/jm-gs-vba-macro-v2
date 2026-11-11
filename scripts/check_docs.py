"""Check VBA procedure references in docs against build/v8_5_vba (stdlib only).

Fails if:
  * a procedure named in the docs is absent from the exported VBA, or
  * a Public procedure/event handler in the exported VBA is undocumented in the docs.

Usage: python scripts/check_docs.py
"""

import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
DOCS = [ROOT / "VBA_CODEBASE.md", ROOT / "VBA_CODE_FLOW.md"]
SRC = ROOT / "build/v8_5_vba"

COMMON = {
    # VBA language / runtime objects commonly mentioned in prose, not procedures to declare.
    "array", "createobject", "getobject", "msgbox", "inputbox", "format", "format$",
    "trim", "trim$", "lcase", "ucase", "cstr", "clng", "cdbl", "cbool", "csng", "cint",
    "isnumeric", "iserror", "isempty", "isdate", "ist missing", "instr", "replace",
    "split", "join", "left", "left$", "right", "right$", "mid", "mid$", "len", "ltrim",
    "rtrim", "space", "string", "string$", "strreverse", "hex", "oct", "abs", "sgn",
    "int", "fix", "round", "rnd", "randomize", "now", "date", "time", "timer", "day",
    "month", "year", "weekday", "monthname", "weekdayname", "dateserial", "datevalue",
    "timevalue", "chron", "environ", "dir", "rgb", "qbcolor", "choose", "switch", "iif",
    "eval", "execute", "scripting.dictionary", "dictionary", "collection", "textstream",
    "filesystemobject", "err", "error", "debug", "shell", "command", "app", "screen",
    "activeworkbook", "activeworksheet", "activestate", "selection", "application",
    "thisworkbook", "workbook", "worksheet", "worksheets", "range", "ranges", "cells",
    "rows", "columns", "column", "row", "name", "value", "text", "count", "add",
    "remove", "removeall", "exists", "item", "keys", "items", "mode", "compare",
    "reset", "exit", "end", "sub", "function", "property", "get", "let", "set",
    "public", "private", "friend", "static", "dim", "set", "let", "set", "call",
    "on error", "resume next", "goto", "if", "then", "else", "elseif", "for", "each",
    "in", "to", "step", "do", "while", "until", "select", "case", "with", "end",
    "with", "variant", "boolean", "integer", "long", "single", "double", "currency",
    "decimal", "date", "string", "object", "bool", "byte", "object",
}

IDENT = re.compile(r"`([A-Za-z_][A-Za-z0-9_]*)`")
PLAIN_IDENT = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\b")


def declared_procedures(src):
    names = set()
    for path in sorted(src.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in (".bas", ".txt", ".cls", ".frm"):
            continue
        text = path.read_text(encoding="utf-8", errors="replace").replace("\r\n", "\n")
        for m in re.finditer(
            r"(?:Public|Private|Friend|Static)?\s*(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\b",
            text,
            re.I,
        ):
            names.add(m.group(2).lower())
    return names


def read_docs_text():
    return "\n".join(p.read_text(encoding="utf-8", errors="replace").replace("\r\n", "\n") for p in DOCS)


def documented_names(text):
    names = set()
    for m in IDENT.finditer(text):
        end = m.end()
        if end < len(text) and text[end] == "(":
            names.add(m.group(1).lower())
    # Dotted calls such as modEngine.SyncPJToCDJ( and sheet event references.
    for m in re.finditer(r"\b(?:Sheet\d+|ThisWorkbook|mod\w+)\.([A-Za-z_]\w*)\s*\(", text):
        names.add(m.group(1).lower())
    return names - COMMON


def public_procedures(src):
    names = set()
    for path in sorted(src.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in (".bas", ".txt", ".cls", ".frm"):
            continue
        text = path.read_text(encoding="utf-8", errors="replace").replace("\r\n", "\n")
        for m in re.finditer(
            r"\bPublic\s+(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\b",
            text,
            re.I,
        ):
            names.add(m.group(2).lower())
        # Event handlers are Private but part of the documented surface.
        for m in re.finditer(
            r"\bPrivate\s+Sub\s+(Worksheet_Change|Worksheet_Activate|Workbook_Open)\b",
            text,
            re.I,
        ):
            names.add(m.group(1).lower())
    return names


def main():
    declared = declared_procedures(SRC)
    docs_text = read_docs_text()
    documented = documented_names(docs_text)
    public = public_procedures(SRC)

    missing = sorted(n for n in documented if n not in declared)
    undocumented = sorted(n for n in public if n not in documented)

    failures = 0
    if missing:
        print("Documented procedures absent from exported VBA:")
        for n in missing:
            print(f"  {n}")
        failures += len(missing)
    if undocumented:
        print("Public procedures/event handlers not referenced in docs:")
        for n in undocumented:
            print(f"  {n}")
        failures += len(undocumented)
    print(f"check_docs: {failures} failure(s); documented={len(documented)}, declared_public={len(public)}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
