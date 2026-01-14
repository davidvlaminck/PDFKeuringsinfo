"""Convenience runner for PyCharm.

This runs the same operation as the shell command:

  python "BTV MIV Keuringsverslagen WVB/detect_btv.py" \
    --update-excel "BTV MIV Keuringsverslagen WVB/Lijst miv.xlsx" \
    --pdf-dir "BTV MIV Keuringsverslagen WVB/pdfs" \
    --excel-out "BTV MIV Keuringsverslagen WVB/Lijst miv.filled.xlsx" \
    --max-pages 2

but is easy to run from within PyCharm.

Notes:
- Configure the run interpreter to use the project venv (.venv).
- To mimic `| head -n 40 | cat`, we only print the first 40 lines of combined stdout/stderr.
"""

from __future__ import annotations

import os
import subprocess
import sys


def main() -> int:
    repo_root = os.path.dirname(os.path.abspath(__file__))

    script_path = os.path.join(repo_root, "BTV MIV Keuringsverslagen WVB", "detect_btv.py")
    update_excel = os.path.join(repo_root, "BTV MIV Keuringsverslagen WVB", "Lijst miv.xlsx")
    pdf_dir = os.path.join(repo_root, "BTV MIV Keuringsverslagen WVB", "pdfs")
    excel_out = os.path.join(repo_root, "BTV MIV Keuringsverslagen WVB", "Lijst miv.filled.xlsx")

    cmd = [
        sys.executable,
        script_path,
        "--update-excel",
        update_excel,
        "--pdf-dir",
        pdf_dir,
        "--excel-out",
        excel_out,
        "--max-pages",
        "2",
    ]

    proc = subprocess.run(
        cmd,
        cwd=repo_root,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )

    lines = (proc.stdout or "").splitlines()
    for line in lines[:40]:
        print(line)

    return proc.returncode


if __name__ == "__main__":
    raise SystemExit(main())
