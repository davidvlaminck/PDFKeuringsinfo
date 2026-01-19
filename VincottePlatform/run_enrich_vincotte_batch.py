"""Convenience runner for PyCharm.

Run this from PyCharm to enrich *all* Excel files in this folder that do NOT
contain 'enriched' in the filename.

It runs `enrich_reports.py --all` and prints only the first 40 lines of output
(similar to piping through `head -n 40`).
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> int:
    vp_dir = Path(__file__).resolve().parent
    script_path = vp_dir / "enrich_reports.py"

    if not script_path.exists():
        print(f"ERROR: not found: {script_path}")
        return 2

    cmd = [sys.executable, str(script_path), "--all"]

    proc = subprocess.run(
        cmd,
        cwd=str(vp_dir),
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )

    for line in (proc.stdout or "").splitlines()[:40]:
        print(line)

    return proc.returncode


if __name__ == "__main__":
    raise SystemExit(main())
