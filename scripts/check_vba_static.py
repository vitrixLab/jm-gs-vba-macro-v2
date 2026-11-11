"""Static VBA checks for build/v8_5_vba (stdlib only).

Flags per procedure:
  * duplicate Dim names in the same procedure (C-01 class)
  * modules missing Option Explicit
  * On Error Resume Next without a following On Error GoTo 0/-1 in the same procedure
Usage: python scripts/check_vba_static.py [root_dir]
Defaults to build/v8_5_vba under the repository root.
"""

import pathlib
import re
import sys

PROC_START = re.compile(
    r"^\s*(?:Public|Private|Friend|Static)?\s*"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\b",
    re.I,
)
PROC_END = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)
DIM_LINE = re.compile(r"^\s*Dim\b(.+)", re.I)
RESUME_NEXT = re.compile(r"^\s*On Error\s+Resume\s+Next\b", re.I)
RESET_ERROR = re.compile(r"^\s*On Error\s+GoTo\s+(0|-1)\b", re.I)
IDENT = re.compile(r"([A-Za-z_]\w*)")


def read_text(path):
    return path.read_text(encoding="utf-8", errors="replace").replace("\r\n", "\n").split("\n")


def procedure_spans(lines):
    spans = []
    i = 0
    while i < len(lines):
        m = PROC_START.match(lines[i])
        if m:
            kind = m.group(1)
            name = m.group(2)
            start = i
            i += 1
            while i < len(lines):
                if PROC_END.match(lines[i]):
                    # Match any End Sub/Function/Property; VBA has no nested procedures.
                    spans.append((start, i, name.lower()))
                    i += 1
                    break
                i += 1
            else:
                spans.append((start, len(lines) - 1, name.lower()))
        else:
            i += 1
    return spans


def dim_names_in_span(lines, start, end):
    names = []
    for idx in range(start + 1, end):
        m = DIM_LINE.match(lines[idx])
        if not m:
            continue
        for segment in m.group(1).split(","):
            seg = segment.strip()
            # Remove array suffixes and attributes, keep the name before As/(
            seg = re.split(r"\s+As\b|\(", seg, maxsplit=1)[0].strip()
            seg = IDENT.match(seg)
            if seg:
                names.append(seg.group(1).lower())
    return names


def check_module(path):
    lines = read_text(path)
    text = "\n".join(lines)
    findings = []
    spans = procedure_spans(lines)
    for start, end, name in spans:
        dims = dim_names_in_span(lines, start, end)
        seen = {}
        for n in dims:
            if n in seen:
                findings.append(f"{path}:{start+1}: duplicate Dim name '{n}' in procedure '{name}' (first seen {seen[n]+1})")
            else:
                seen[n] = idx = start  # not used for line
        # unreset On Error Resume Next
        for idx in range(start + 1, end):
            if RESUME_NEXT.match(lines[idx]):
                if not any(RESET_ERROR.match(lines[j]) for j in range(idx + 1, end)):
                    findings.append(f"{path}:{idx+1}: 'On Error Resume Next' not reset before end of procedure '{name}'")
    if "option explicit" not in text.lower():
        findings.append(f"{path}: module missing 'Option Explicit'")
    return findings


def main():
    root = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else pathlib.Path("build/v8_5_vba")
    total = 0
    for path in sorted(root.rglob("*")):
        if path.is_file() and path.suffix.lower() in (".bas", ".txt", ".cls", ".frm"):
            findings = check_module(path)
            total += len(findings)
            for f in findings:
                print(f)
    print(f"check_vba_static: {total} finding(s) in {root}")
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
