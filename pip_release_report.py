#!/usr/bin/env python3
import argparse
import datetime as dt
import json
import re
from pathlib import Path

import requests

try:
    from packaging.version import parse as vparse
except Exception:  # pragma: no cover - best-effort import
    vparse = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
except Exception as exc:  # pragma: no cover - best-effort import
    raise SystemExit(
        "openpyxl is required to write the Excel file. "
        "Install it with: pip install openpyxl"
    ) from exc

LINE_RE = re.compile(r"^\s*([A-Za-z0-9_.-]+)\s*==\s*([^\s;]+)")
NPM_VERSION_RE = re.compile(r"(\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?)")
PYPI_URL = "https://pypi.org/pypi/{name}/json"
NPM_URL = "https://registry.npmjs.org/{name}"
EXCLUDED_PACKAGES = {
    "colorlog",
    "django-extensions",
    "redis",
    "setuptools",
    "wheel",
}


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


def parse_package_json(path: Path) -> list[tuple[str, str]]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []
    items: list[tuple[str, str]] = []
    deps = data.get("dependencies", {})
    if not isinstance(deps, dict):
        return items
    for name, spec in deps.items():
        if not isinstance(spec, str):
            continue
        if name.startswith("@angular/"):
            continue
        items.append((name, spec))
    return items


def extract_npm_version(spec: str) -> str | None:
    match = NPM_VERSION_RE.search(spec)
    return match.group(1) if match else None


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


def is_stable_version(version: str) -> bool:
    if not vparse:
        return False
    parsed = vparse(version)
    # Exclude pre, post, dev, and local versions
    return not (parsed.is_prerelease or parsed.is_postrelease or parsed.is_devrelease or parsed.local)


def get_latest_version(info: dict, releases: dict) -> str | None:
    if not vparse:
        return info.get("version")
    versions = [v for v in releases.keys() if v and is_stable_version(v)]
    if not versions:
        return None
    return str(max(versions, key=vparse))


def get_latest_version_same_major(
    releases: dict, current_version: str
) -> str | None:
    if not vparse:
        return None
    try:
        current_major = vparse(current_version).release[0]
    except Exception:
        return None
    candidates = []
    for version in releases.keys():
        if not version or not is_stable_version(version):
            continue
        parsed = vparse(version)
        if parsed.release and parsed.release[0] == current_major:
            candidates.append(version)
    if not candidates:
        return None
    return str(max(candidates, key=vparse))


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


def fetch_npm(name: str, session: requests.Session, cache: dict) -> dict | None:
    if name in cache:
        return cache[name]
    try:
        encoded = requests.utils.quote(name, safe="@/")
        resp = session.get(NPM_URL.format(name=encoded), timeout=20)
        if resp.status_code != 200:
            cache[name] = None
            return None
        cache[name] = resp.json()
        return cache[name]
    except requests.RequestException:
        cache[name] = None
        return None


def get_npm_release_date(time_map: dict, version: str) -> dt.date | None:
    if not version:
        return None
    ts = time_map.get(version)
    if not ts:
        return None
    try:
        return dt.datetime.fromisoformat(ts.replace("Z", "+00:00")).date()
    except ValueError:
        return None


def get_npm_latest_version(data: dict) -> str | None:
    dist_tags = data.get("dist-tags", {}) if isinstance(data, dict) else {}
    latest = dist_tags.get("latest")
    if not vparse:
        return latest if isinstance(latest, str) else None
    if isinstance(latest, str) and is_stable_version(latest):
        return latest
    time_map = data.get("time", {}) if isinstance(data, dict) else {}
    versions = [
        v
        for v in time_map.keys()
        if v not in ("created", "modified") and is_stable_version(v)
    ]
    if not versions:
        return None
    return str(max(versions, key=vparse))


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate an Excel report comparing pinned versions to latest PyPI/NPM releases."
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
    npm_files = list(root.rglob("package.json"))
    if not pip_files and not npm_files:
        print(f"No .pip or package.json files found under {root}")
        return 1

    grouped: dict[str, list[tuple[str, str, str]]] = {}
    for path in pip_files:
        folder = path.parent.name or path.parent.as_posix()
        grouped.setdefault(folder, [])
        grouped[folder].extend((name, version, "pypi") for name, version in parse_pip_file(path))
    for path in npm_files:
        folder = path.parent.name or path.parent.as_posix()
        grouped.setdefault(folder, [])
        grouped[folder].extend((name, version, "npm") for name, version in parse_package_json(path))

    wb = Workbook()
    wb.remove(wb.active)

    session = requests.Session()
    session.headers.update({"User-Agent": "pip-release-report/1.0"})
    pypi_cache: dict[str, dict | None] = {}
    npm_cache: dict[str, dict | None] = {}

    headers = [
        "package",
        "current_version",
        "current_release_date",
        "latest_version",
        "latest_release_date",
        "days_difference",
        "days_since_latest_release",
    ]
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True, size=16)
    header_alignment = Alignment(horizontal="center", vertical="center")
    body_font = Font(size=14)
    body_alignment = Alignment(horizontal="center", vertical="center")
    package_alignment = Alignment(horizontal="left", vertical="center")
    even_fill = PatternFill("solid", fgColor="F2F6FA")
    odd_fill = PatternFill("solid", fgColor="FFFFFF")
    alert_fill = PatternFill("solid", fgColor="F8D7DA")
    warning_fill = PatternFill("solid", fgColor="FFF3CD")
    alert_threshold_days = 2 * 365 - 62

    today = dt.date.today()
    for folder, items in sorted(grouped.items()):
        ws = wb.create_sheet(title=sanitize_sheet_name(folder))
        ws.append(headers)
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
        rows: list[list] = []
        for name, current_version, ecosystem in items:
            if name.lower() in EXCLUDED_PACKAGES:
                continue
            if ecosystem == "pypi":
                data = fetch_pypi(name, session, pypi_cache)
                if not data:
                    rows.append([name, current_version, None, None, None, None, None])
                    continue

                info = data.get("info", {})
                releases = data.get("releases", {})
                current_date = get_release_date(releases, current_version)
                if name.lower() == "django":
                    latest_version = get_latest_version_same_major(releases, current_version)
                    if not latest_version:
                        latest_version = get_latest_version(info, releases)
                else:
                    latest_version = get_latest_version(info, releases)
                latest_date = get_release_date(releases, latest_version) if latest_version else None
            else:
                data = fetch_npm(name, session, npm_cache)
                if not data:
                    rows.append([name, current_version, None, None, None, None, None])
                    continue

                time_map = data.get("time", {})
                current_resolved = extract_npm_version(current_version)
                current_date = get_npm_release_date(time_map, current_resolved)
                latest_version = get_npm_latest_version(data)
                latest_date = get_npm_release_date(time_map, latest_version)

            days_diff = None
            if current_date and latest_date:
                days_diff = (latest_date - current_date).days
            days_since_latest = None
            if latest_date:
                days_since_latest = (today - latest_date).days

            rows.append(
                [
                    name,
                    current_version,
                    current_date if current_date else None,
                    latest_version,
                    latest_date if latest_date else None,
                    days_diff,
                    days_since_latest,
                ]
            )
        rows.sort(
            key=lambda r: (r[5] is None, r[5] if r[5] is not None else -1),
            reverse=True,
        )
        for row in rows:
            ws.append(row)
        for row in range(2, ws.max_row + 1):
            days_value = ws.cell(row=row, column=6).value
            if isinstance(days_value, int) and days_value > alert_threshold_days:
                fill = alert_fill
            else:
                fill = even_fill if row % 2 == 0 else odd_fill
            for col in range(1, len(headers) + 1):
                cell = ws.cell(row=row, column=col)
                cell.fill = fill
                cell.font = body_font
                cell.alignment = package_alignment if col == 1 else body_alignment
                if col in (3, 5) and cell.value:
                    cell.number_format = "DD-MMM-YYYY"
                if col == 7 and isinstance(cell.value, int) and cell.value > alert_threshold_days:
                    cell.fill = warning_fill
        for col in range(1, len(headers) + 1):
            max_len = 0
            for row in range(1, ws.max_row + 1):
                value = ws.cell(row=row, column=col).value
                if value is None:
                    continue
                max_len = max(max_len, len(str(value)))
            header_len = len(str(headers[col - 1]))
            width = max(max_len + 2, int(header_len * 1.25) + 4)
            ws.column_dimensions[chr(64 + col)].width = width

    wb.save(args.output)
    print(f"Wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
