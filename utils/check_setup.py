"""Small setup check for uv-based environments."""

from __future__ import annotations

import importlib.util

REQUIRED = [
    "requests",
    "pandas",
    "openpyxl",
    "pytz",
    "tqdm",
    "shapely",
    "jwt",
    "msal",
]

OPTIONAL = [
    "geopandas",
]


def _is_installed(module_name: str) -> bool:
    return importlib.util.find_spec(module_name) is not None


def main() -> int:
    missing_required = [name for name in REQUIRED if not _is_installed(name)]
    missing_optional = [name for name in OPTIONAL if not _is_installed(name)]

    if missing_required:
        print("Missing required modules:", ", ".join(missing_required))
        print("Run: uv sync")
        return 1

    print("Required modules: OK")
    if missing_optional:
        print("Optional modules missing:", ", ".join(missing_optional))
        print("Run: uv sync --extra geo")
    else:
        print("Optional modules: OK")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
