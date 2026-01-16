#!/usr/bin/env python3
import argparse
import datetime as dt
import re
from pathlib import Path

import requests

try:
    from packaging.version import parse as vparse
except Exception:  # pragma: no cover - best-effort import
    vparse = None

try:
    from openpyxl import Workbook
except Exception as exc:  # pragma: no cover - best-effort import
    raise SystemExit(
        "openpyxl is required to write the Excel file. "
        "Install it with: pip install openpyxl"
    ) from exc

LINE_RE = re.compile(r"^\s*([A-Za-z0-9_.-]+)\s*==\s*([^\s;]+)")
PYPI_URL = "https://pypi.org/pypi/{name}/json"


def parse_pip_file(path: Path) -> list[tuple[str, str]]:
    items: list[tuple[str, str]] = []
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        match = LINE_RE.match(line)
        if not match:
            continue
        name, version = match.group(1), match.group(2)
        items.append((name, version))
    return items


def get_release_date(releases: dict, version: str) -> dt.date | None:
    files = releases.get(version) or []
    dates = []
    for entry in files:
        ts = entry.get("upload_time_iso_8601") or entry.get("upload_time")
        if not ts:
            continue
        try:
            dates.append(dt.datetime.fromisoformat(ts.replace("Z", "+00:00")))
        except ValueError:
            continue
    if not dates:
        return None
    return min(dates).date()


def get_latest_version(info: dict, releases: dict) -> str | None:
    latest = info.get("version")
    if latest:
        return latest
    if not vparse:
        return None
    versions = [v for v in releases.keys() if v]
    if not versions:
        return None
    return str(max(versions, key=vparse))


def sanitize_sheet_name(name: str) -> str:
    cleaned = re.sub(r"[\\/*?:\\[\\]]", "_", name)
    return cleaned[:31] if cleaned else "Sheet"


def fetch_pypi(name: str, session: requests.Session, cache: dict) -> dict | None:
    if name in cache:
        return cache[name]
    try:
        resp = session.get(PYPI_URL.format(name=name), timeout=20)
        if resp.status_code != 200:
            cache[name] = None
            return None
        cache[name] = resp.json()
        return cache[name]
    except requests.RequestException:
        cache[name] = None
        return None


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate an Excel report comparing pinned versions to latest PyPI releases."
    )
    parser.add_argument(
        "--root",
        default=".",
        help="Root directory to scan for .pip files (default: current directory).",
    )
    parser.add_argument(
        "--output",
        default="pip_release_report.xlsx",
        help="Output Excel file path (default: pip_release_report.xlsx).",
    )
    args = parser.parse_args()

    root = Path(args.root).resolve()
    pip_files = list(root.rglob("*.pip"))
    if not pip_files:
        print(f"No .pip files found under {root}")
        return 1

    grouped: dict[str, list[tuple[str, str]]] = {}
    for path in pip_files:
        folder = path.parent.name or path.parent.as_posix()
        grouped.setdefault(folder, [])
        grouped[folder].extend(parse_pip_file(path))

    wb = Workbook()
    wb.remove(wb.active)

    session = requests.Session()
    session.headers.update({"User-Agent": "pip-release-report/1.0"})
    cache: dict[str, dict | None] = {}

    headers = [
        "package",
        "current_version",
        "current_release_date",
        "latest_version",
        "latest_release_date",
        "days_difference",
    ]

    for folder, items in sorted(grouped.items()):
        ws = wb.create_sheet(title=sanitize_sheet_name(folder))
        ws.append(headers)
        for name, current_version in items:
            data = fetch_pypi(name, session, cache)
            if not data:
                ws.append([name, current_version, None, None, None, None])
                continue

            info = data.get("info", {})
            releases = data.get("releases", {})
            current_date = get_release_date(releases, current_version)
            latest_version = get_latest_version(info, releases)
            latest_date = get_release_date(releases, latest_version) if latest_version else None

            days_diff = None
            if current_date and latest_date:
                days_diff = (latest_date - current_date).days

            ws.append(
                [
                    name,
                    current_version,
                    current_date.isoformat() if current_date else None,
                    latest_version,
                    latest_date.isoformat() if latest_date else None,
                    days_diff,
                ]
            )

    wb.save(args.output)
    print(f"Wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
