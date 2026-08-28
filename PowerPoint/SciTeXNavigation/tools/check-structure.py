#!/usr/bin/env python3
"""Check an exported .bas for the structural errors the VBA editor finds late.

WHY THIS EXISTS
---------------
Importing a module into PowerPoint is the only real compiler available here,
and it needs AccessVBOM, a running Windows host and a file the operator has
open.  That round trip is slow enough that an unbalanced ``If`` is discovered
after the import rather than before it.  These checks read the text.

They are deliberately narrow.  This is not a VBA parser and will not catch a
type error; it catches the mistakes that come from editing a long module with
a script -- a block left open, a routine ended with the wrong keyword, a name
called but never defined.

READING VBA AS TEXT
-------------------
Three things make naive line matching wrong, and all three produced false
findings before they were handled:

  comments      ``' Next slide`` is prose, not a ``Next``.
  labels        ``NextSlide:`` starts with ``Next``.
  continuations ``If a = 1 And _`` / ``   b = 2 Then`` is ONE statement.  The
                physical line does not end in ``Then``, so a single-line-If
                test misfires on it.

So: join continuations, strip comments and string literals, then count.

EXIT CODES
----------
0   no findings
10  findings
2   usage / unreadable
"""
from __future__ import annotations

import argparse
import pathlib
import re
import sys

EXIT_CLEAN, EXIT_FINDINGS, EXIT_USAGE = 0, 10, 2

STRING = re.compile(r'"[^"]*"')
LABEL = re.compile(r"^[A-Za-z_]\w*:\s*$")
OPENERS = {
    "Sub": ("End Sub",),
    "Function": ("End Function",),
    "Property": ("End Property",),
    "With": ("End With",),
    "Type": ("End Type",),
}


def logical_lines(text: str):
    """Physical lines joined at ``_`` continuations, comments and strings gone.

    Yields (line_number_of_first_physical_line, cleaned_text).
    """
    pending, start = "", None
    for number, raw in enumerate(text.splitlines(), start=1):
        line = raw.rstrip()
        if start is None:
            start = number
        # A quote inside a comment and an apostrophe inside a string are the
        # same character; drop strings first so the comment cut is honest.
        stripped = STRING.sub('""', line)
        cut = stripped.find("'")
        if cut >= 0:
            keep = len(STRING.sub('""', line[:cut]))
            line = line[:keep]
        line = line.rstrip()
        if line.endswith(" _"):
            pending += line[:-1]
            continue
        joined = (pending + line).strip()
        pending = ""
        if joined:
            yield start, joined
        start = None


def check(path: pathlib.Path, findings: list):
    text = path.read_text(encoding="utf-8", errors="replace")
    lines = list(logical_lines(text))

    stack = []
    defined, called = set(), {}
    for number, line in lines:
        if LABEL.match(line):
            continue
        words = line.split()
        head = words[0] if words else ""
        if head in ("Public", "Private", "Friend"):
            words = words[1:]
            head = words[0] if words else ""
        if head == "Static":
            words = words[1:]
            head = words[0] if words else ""

        if head in ("Sub", "Function", "Property") and len(words) > 1:
            name = re.split(r"[ (]", words[1] if head != "Property" else words[2])[0]
            defined.add(name)
            stack.append((head, number))
        elif head in ("With", "Type"):
            stack.append((head, number))
        elif line.startswith("End "):
            want = line.split()[1]
            if not stack:
                findings.append(f"{path.name}:{number}: End {want} with nothing open")
            else:
                opener, opened = stack.pop()
                if opener != want and not (opener == "Property" and want == "Property"):
                    findings.append(
                        f"{path.name}:{number}: End {want} closes {opener} "
                        f"opened at line {opened}")
        elif head == "If":
            # A multi-line If ends in Then; a single-line If has code after it.
            if re.search(r"\bThen\s*$", line):
                stack.append(("If", number))
        elif head in ("For", "Do", "While", "Select"):
            stack.append((head, number))
        elif head in ("Next", "Loop", "Wend"):
            expect = {"Next": "For", "Loop": "Do", "Wend": "While"}[head]
            if not stack:
                findings.append(f"{path.name}:{number}: {head} with nothing open")
            else:
                opener, opened = stack.pop()
                if opener != expect:
                    findings.append(
                        f"{path.name}:{number}: {head} closes {opener} "
                        f"opened at line {opened}")

        for name in re.findall(r"\b([A-Z]\w+)\s*\(", line):
            called.setdefault(name, number)
        for name in re.findall(r"^\s*([A-Z]\w+)\s+[^=]*$", line):
            called.setdefault(name, number)

    for opener, opened in stack:
        findings.append(f"{path.name}:{opened}: {opener} is never closed")

    known = defined | VBA_BUILTINS
    for name, number in sorted(called.items(), key=lambda kv: kv[1]):
        if name not in known and name.startswith(("Set", "Fit", "Rebuild", "Column",
                                                  "Plan", "Layout", "LayOut", "Measure",
                                                  "Belongs", "Entry", "Usable", "Highest",
                                                  "Largest", "Normalise", "Note", "Backup")):
            findings.append(f"{path.name}:{number}: calls {name}, which is not defined here")


VBA_BUILTINS = {
    "If", "For", "Do", "Select", "Case", "Set", "Dim", "Err", "Raise", "RGB", "IIf",
    "CLng", "CStr", "CSng", "Val", "Len", "Mid", "InStr", "InStrRev", "Trim", "Left",
    "Right", "StrComp", "Abs", "Int", "Application", "MsgBox", "Array", "UBound",
    "LBound", "Split", "Join", "Replace", "UCase", "LCase", "Format", "Now", "Date",
    "Time", "Chr", "Asc", "Exit", "End", "Next", "Loop", "Wend", "Then", "Else",
    "ElseIf", "On", "Error", "Resume", "GoTo", "Call", "With", "While", "Until",
    "Not", "And", "Or", "Is", "Nothing", "True", "False", "New", "Me", "Public",
    "Private", "Const", "Type", "Function", "Sub", "Property", "Option", "Attribute",
    "Sqr", "Round", "Space", "String", "Rnd", "Timer", "Environ", "Dir", "Kill",
}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("modules", nargs="+", type=pathlib.Path, help=".bas files")
    args = ap.parse_args()

    findings: list[str] = []
    for path in args.modules:
        if not path.is_file():
            print(f"ERROR: not a file: {path}", file=sys.stderr)
            return EXIT_USAGE
        check(path, findings)

    for line in findings:
        print(line)
    print(f"{len(args.modules)} module(s), {len(findings)} finding(s)")
    return EXIT_FINDINGS if findings else EXIT_CLEAN


if __name__ == "__main__":
    raise SystemExit(main())
