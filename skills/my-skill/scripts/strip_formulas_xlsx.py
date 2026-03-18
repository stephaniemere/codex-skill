#!/usr/bin/env python3
"""Create a values-only copy of an .xlsx file by removing all formula nodes from worksheets."""

from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path

FORMULA_EMPTY_RE = re.compile(r"<f[^>/]*/>")
FORMULA_BLOCK_RE = re.compile(r"<f[^>]*>.*?</f>", re.S)


def strip_formulas(xml_text: str) -> str:
    text = FORMULA_EMPTY_RE.sub("", xml_text)
    text = FORMULA_BLOCK_RE.sub("", text)
    return text


def run(inp: Path, out: Path) -> int:
    with zipfile.ZipFile(inp, "r") as zin, zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                txt = data.decode("utf-8")
                txt = strip_formulas(txt)
                data = txt.encode("utf-8")
            zout.writestr(item, data)
    return 0


def main() -> int:
    p = argparse.ArgumentParser(description="Remove worksheet formulas from an xlsx and keep cached values")
    p.add_argument("--input", required=True, type=Path, help="Input .xlsx")
    p.add_argument("--output", required=True, type=Path, help="Output .xlsx")
    args = p.parse_args()
    return run(args.input, args.output)


if __name__ == "__main__":
    raise SystemExit(main())
