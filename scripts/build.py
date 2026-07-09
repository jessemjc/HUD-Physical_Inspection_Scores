#!/usr/bin/env python3
"""
Rebuilds the FHA Multifamily Inspection Score Lookup tool from HUD's
published datasets.

Downloads:
  - Insured Active Multifamily FHA Property Addresses (address index)
  - REAC Physical Inspection Scores and Release Dates (score records, legacy .xls)

Then re-runs the same extraction/embedding process used to hand-build the
tool, and writes the finished, self-contained HTML file to docs/index.html
so GitHub Pages can serve it.

If HUD reorganizes their site and one of the URLs below 404s, this script
will fail loudly (see the requests.raise_for_status() calls) rather than
silently publishing stale or broken data. Check the "Data & Research" /
"Multifamily Data" pages on hud.gov for the new link and update the
constants below.
"""

import datetime
import json
import subprocess
import sys
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parent.parent
DOWNLOAD_DIR = ROOT / "downloads"
DOWNLOAD_DIR.mkdir(exist_ok=True)

ADDRESS_URL = "https://www.hud.gov/sites/dfiles/Housing/documents/InsuredActiveMultifamilyFHAPropertyAddresses.xlsx"
SCORES_URL = "https://www.hud.gov/sites/default/files/Housing/documents/MF-Inspection-Report.xls"

ADDRESS_FILE = DOWNLOAD_DIR / "addresses.xlsx"
SCORES_FILE_XLS = DOWNLOAD_DIR / "scores.xls"
SCORES_FILE_XLSX = DOWNLOAD_DIR / "scores.xlsx"

TEMPLATE_PATH = ROOT / "templates" / "shell.html"
OUTPUT_PATH = ROOT / "docs" / "index.html"

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; fha-lookup-updater/1.0)"}


def download(url: str, dest: Path) -> None:
    print(f"Downloading {url} -> {dest}")
    resp = requests.get(url, headers=HEADERS, timeout=120)
    resp.raise_for_status()
    dest.write_bytes(resp.content)
    print(f"  saved {len(resp.content):,} bytes")


def convert_legacy_xls(src: Path, dest_dir: Path) -> Path:
    """Convert the legacy .xls scores file to .xlsx using headless LibreOffice."""
    print(f"Converting {src} to .xlsx via LibreOffice headless...")
    subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(dest_dir),
            str(src),
        ],
        check=True,
        capture_output=True,
    )
    converted = dest_dir / (src.stem + ".xlsx")
    if not converted.exists():
        raise RuntimeError(f"Expected converted file at {converted} but it wasn't created")
    return converted


def clean(s):
    if pd.isna(s):
        return ""
    return str(s).replace("-", "").replace(" ", "").strip()


def build_address_records(addr_df: pd.DataFrame) -> list:
    """Returns [fha_clean, property_id, fha_raw, addr_name, addr_city, addr_state] rows.

    addr_name/addr_city/addr_state are kept only as a fallback -- the scores
    spreadsheet's Property Name/City/state_code are the primary source, since
    that sheet also covers properties with no FHA-insured address record.
    """
    addr_df["fha_clean"] = addr_df["fha_number"].apply(clean)

    def strip_or_none(s):
        if pd.isna(s):
            return None
        s = str(s).strip()
        return s if s else None

    records = []
    for _, row in addr_df.iterrows():
        if not row["fha_clean"]:
            continue
        records.append([
            row["fha_clean"],
            row["property_id"] if pd.notna(row["property_id"]) else "",
            row["fha_number"] if pd.notna(row["fha_number"]) else "",
            strip_or_none(row["property_name_text"]),
            strip_or_none(row["city_name_text"]),
            strip_or_none(row["state_code"]),
        ])
    return records


def build_score_records(scores_df: pd.DataFrame) -> dict:
    """Returns property_id -> [name, city, state, [[s1,d1],[s2,d2],[s3,d3]]].

    Name/city/state come straight from the HUD inspection scores spreadsheet
    (Property Name / City / state_code), which is the primary source for
    project info -- it covers properties even when they have no FHA-insured
    address record.
    """
    def strip_or_none(s):
        if pd.isna(s):
            return None
        s = str(s).strip()
        return s if s else None

    records = {}
    for _, row in scores_df.iterrows():
        pid = row["REMS Property Id"]
        if pd.isna(pid):
            continue
        pairs = []
        for i in (1, 2, 3):
            sc = row.get(f"Inspection Score{i}")
            dt = row.get(f"Release Date {i}")
            pairs.append([
                sc if pd.notna(sc) else None,
                dt if pd.notna(dt) else None,
            ])
        records[pid] = [
            strip_or_none(row.get("Property Name")),
            strip_or_none(row.get("City")),
            strip_or_none(row.get("state_code")),
            pairs,
        ]
    return records


def main():
    # 1. Download source files
    download(ADDRESS_URL, ADDRESS_FILE)
    download(SCORES_URL, SCORES_FILE_XLS)

    # 2. Convert legacy .xls scores file
    converted = convert_legacy_xls(SCORES_FILE_XLS, DOWNLOAD_DIR)
    converted.rename(SCORES_FILE_XLSX)

    # 3. Load both spreadsheets
    print("Reading address spreadsheet...")
    addr_df = pd.read_excel(ADDRESS_FILE, dtype=str)
    print(f"  {len(addr_df):,} address rows")

    print("Reading scores spreadsheet...")
    scores_df = pd.read_excel(SCORES_FILE_XLSX, dtype=str)
    print(f"  {len(scores_df):,} score rows")

    # 4. Sanity-check expected columns exist -- fail loudly if HUD changes their schema
    required_addr_cols = {
        "property_name_text", "property_id", "fha_number",
        "city_name_text", "state_code",
    }
    missing_addr = required_addr_cols - set(addr_df.columns)
    if missing_addr:
        sys.exit(f"ERROR: address spreadsheet is missing expected columns: {missing_addr}")

    required_score_cols = {
        "REMS Property Id", "Inspection Score1", "Release Date 1",
        "Property Name", "City", "state_code",
    }
    missing_score = required_score_cols - set(scores_df.columns)
    if missing_score:
        sys.exit(f"ERROR: scores spreadsheet is missing expected columns: {missing_score}")

    # 5. Extract compact records
    addr_records = build_address_records(addr_df)
    score_records = build_score_records(scores_df)
    print(f"Extracted {len(addr_records):,} address records, {len(score_records):,} score records")

    # 6. Embed into the HTML template
    template = TEMPLATE_PATH.read_text()
    addr_json = json.dumps(addr_records, separators=(",", ":"))
    score_json = json.dumps(score_records, separators=(",", ":"))

    html = template
    html = html.replace("__ADDR_JSON__", addr_json)
    html = html.replace("__SCORE_JSON__", score_json)
    html = html.replace("__ADDR_COUNT__", f"{len(addr_records):,}")
    html = html.replace("__SCORE_COUNT__", f"{len(score_records):,}")
    html = html.replace(
        "__DATA_DATE__",
        "Snapshot: " + datetime.date.today().strftime("%B %d, %Y"),
    )

    OUTPUT_PATH.parent.mkdir(exist_ok=True)
    OUTPUT_PATH.write_text(html)
    print(f"Wrote {OUTPUT_PATH} ({len(html) / 1024 / 1024:.2f} MB)")


if __name__ == "__main__":
    main()
