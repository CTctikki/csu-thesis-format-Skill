#!/usr/bin/env python3
"""Smoke test for the specialized Word structure audit script."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "audit_csu_word_structures.py"
TEMPLATE = ROOT / "assets" / "csu-thesis-template.docx"


def main() -> int:
    if not SCRIPT.exists():
        raise AssertionError(f"missing script: {SCRIPT}")
    if not TEMPLATE.exists():
        raise AssertionError(f"missing template: {TEMPLATE}")

    result = subprocess.run(
        [sys.executable, str(SCRIPT), str(TEMPLATE)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if result.returncode != 0:
        raise AssertionError(result.stderr or result.stdout)

    output = result.stdout
    for token in [
        "# CSU Thesis Word",
        "TOC",
        "word/header",
        "OMath",
    ]:
        if token not in output:
            raise AssertionError(f"missing output token: {token}")

    print("PASS test_audit_csu_word_structures")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
