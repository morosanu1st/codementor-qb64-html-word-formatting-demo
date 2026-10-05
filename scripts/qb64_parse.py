"""Shared helpers: split a QB64 source file into STEP sections."""
from __future__ import annotations

import re
from dataclasses import dataclass, field

STEP_RE = re.compile(r"^\s*'\s*STEP\s+(\d+)\s*:\s*(.+?)\s*$", re.IGNORECASE)

KEYWORDS = {
    "AS", "CASE", "CLS", "DATA", "DIM", "DO", "ELSE", "ELSEIF", "END", "FOR", "FUNCTION",
    "GOSUB", "GOTO", "IF", "INPUT", "INTEGER", "IS", "LET", "LINE", "LONG", "LOOP", "NEXT",
    "PRINT", "READ", "RETURN", "SELECT", "SHARED", "SINGLE", "DOUBLE", "STEP", "STRING",
    "SUB", "THEN", "TO", "UNTIL", "USING", "WEND", "WHILE", "AND", "OR", "NOT", "MOD",
}


@dataclass
class Step:
    number: int
    title: str
    lines: list[str] = field(default_factory=list)

    @property
    def code(self) -> str:
        # Trim blank lines at both ends but keep inner spacing/indentation intact.
        lines = list(self.lines)
        while lines and not lines[0].strip():
            lines.pop(0)
        while lines and not lines[-1].strip():
            lines.pop()
        return "\n".join(lines)


def parse_steps(source: str) -> list[Step]:
    steps: list[Step] = []
    current: Step | None = None
    for raw in source.splitlines():
        line = raw.rstrip().replace("\t", "    ")
        m = STEP_RE.match(line)
        if m:
            current = Step(int(m.group(1)), m.group(2))
            steps.append(current)
        elif current is not None:
            current.lines.append(line)
    return steps


def program_title(source: str, default: str = "QB64 Program") -> str:
    """Use the first non-banner comment line as the title."""
    for raw in source.splitlines():
        text = raw.strip()
        if text.startswith("'"):
            text = text.lstrip("'").strip()
            if text and not set(text) <= set("=-*"):
                return text.split(" - ")[0]
    return default
