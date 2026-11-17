"""Inspect a DOCX produced by the Markdown-to-MathType workflow.

The verifier is deliberately structural: it does not pretend that a raster
preview proves editability.  Instead it checks that each formula is an
embedded Equation.DSMT4 OLE object, that no native Word OMML remains, and that
each MathType payload carries the requested Times New Roman primary font.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path


OMML = re.compile(rb"<m:oMath(?:\s|>)")
PROG_ID = b'ProgID="Equation.DSMT4"'
TNR = b"Times New Roman"
CAMBRIA_MATH = b"Cambria Math"


def inspect(path: Path) -> dict[str, object]:
    if not path.is_file():
        raise FileNotFoundError(f"DOCX was not found: {path}")

    with zipfile.ZipFile(path) as package:
        corrupt_member = package.testzip()
        if corrupt_member:
            raise ValueError(f"DOCX ZIP CRC failed in: {corrupt_member}")
        try:
            document = package.read("word/document.xml")
        except KeyError as exc:
            raise ValueError("word/document.xml is missing") from exc
        try:
            styles = package.read("word/styles.xml")
        except KeyError as exc:
            raise ValueError("word/styles.xml is missing") from exc

        ole_parts = sorted(
            item.filename
            for item in package.infolist()
            if item.filename.startswith("word/embeddings/oleObject")
            and item.filename.endswith(".bin")
        )
        payloads = [package.read(name) for name in ole_parts]

    return {
        "docx": str(path.resolve()),
        "zip_crc": "pass",
        "omml": len(OMML.findall(document)),
        "equation_dsm_t4": document.count(PROG_ID),
        "ole_objects": len(ole_parts),
        "ole_with_times_new_roman": sum(TNR in payload for payload in payloads),
        "ole_with_cambria_math": sum(CAMBRIA_MATH in payload for payload in payloads),
        "body_styles_times_new_roman_11": bool(
            TNR in styles and b'<w:sz w:val="22"' in styles
        ),
    }


def enforce(info: dict[str, object], stage: str, expected: int | None) -> None:
    omml = int(info["omml"])
    ole = int(info["ole_objects"])
    prog = int(info["equation_dsm_t4"])
    tnr = int(info["ole_with_times_new_roman"])
    cambria = int(info["ole_with_cambria_math"])
    body_tnr_11 = bool(info["body_styles_times_new_roman_11"])

    if stage in {"omml", "mathtype"} and not body_tnr_11:
        raise ValueError("The Word text/style layer does not declare Times New Roman at 11 pt.")

    if stage == "omml":
        if omml < 1:
            raise ValueError("The intermediate DOCX contains no OMML equations.")
        if ole or prog:
            raise ValueError("The OMML intermediate unexpectedly already contains MathType OLE objects.")
    elif stage == "mathtype":
        if omml:
            raise ValueError(f"{omml} native Word OMML equations remain after conversion.")
        if ole < 1 or prog < 1:
            raise ValueError("No Equation.DSMT4 MathType OLE equations were found.")
        if ole != prog:
            raise ValueError(f"OLE count ({ole}) and Equation.DSMT4 count ({prog}) differ.")
        if expected is not None and ole != expected:
            raise ValueError(f"Expected {expected} MathType equations, found {ole}.")
        if tnr != ole:
            raise ValueError(f"Times New Roman was not recorded in every MathType payload ({tnr}/{ole}).")
        if cambria:
            raise ValueError(f"Cambria Math remains in {cambria} MathType payload(s).")
    elif stage != "inspect":
        raise ValueError(f"Unknown stage: {stage}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("--stage", choices=("inspect", "omml", "mathtype"), default="inspect")
    parser.add_argument("--expected-equations", type=int)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    try:
        info = inspect(args.docx)
        enforce(info, args.stage, args.expected_equations)
    except Exception as exc:  # clear terminal error for BAT/PowerShell callers
        print(f"VERIFY FAILED: {exc}", file=sys.stderr)
        raise SystemExit(2) from exc

    if args.json:
        print(json.dumps(info, ensure_ascii=False))
    else:
        for key, value in info.items():
            print(f"{key}: {value}")


if __name__ == "__main__":
    main()
