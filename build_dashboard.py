from __future__ import annotations

import csv
import json
import os
import re
from collections import Counter
from pathlib import Path


SOURCE_CSV = Path("DynamicList.CSV")
OUTPUT_FILES = (Path("index.html"), Path("dashboard.html"))
DATA_FILE = Path("data.json")
SALES_SOURCE_JSON = Path(os.getenv("THREE_W_DATA_JSON", "../-3W bkg dashboard/dist/data.json"))
SOURCE_ENCODINGS = ("cp949", "utf-8-sig", "utf-8")

EFC_EXCLUDED_ORIGINS = {"", "CN", "KR", "JP", "US"}
SOUTH_EAST_ASIA = {"TH", "VN", "ID", "MY", "SG", "PH", "KH", "MM"}
ISC = {"IN", "PK", "LK", "BD"}
MIDDLE_EAST = {"AE", "OM", "SA", "BH", "KW"}
RED_SEA_COUNTRIES = {"EG", "JO"}
RED_SEA_PORTS = {"JED", "AQJ", "SKN"}
AFRICA = {"TZ", "KE"}


def clean(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def number(value: object) -> float:
    text = clean(value).replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def whole(value: float) -> int | float:
    return int(value) if float(value).is_integer() else round(value, 2)


def key_code(value: object) -> str:
    return clean(value).upper()


def normalize_salesperson(value: object) -> str:
    text = clean(value)
    if not text:
        return ""
    names: list[str] = []
    seen: set[str] = set()
    for name in re.split(r"[,;]+", text):
        normalized = clean(name).upper()
        if normalized and normalized not in seen:
            seen.add(normalized)
            names.append(normalized)
    return ", ".join(names)


def load_salesperson_lookup() -> tuple[dict[str, str], dict]:
    meta = {
        "source": str(SALES_SOURCE_JSON),
        "found": SALES_SOURCE_JSON.exists(),
        "sourceRows": 0,
        "mappingKeys": 0,
        "duplicateKeys": 0,
    }
    if not SALES_SOURCE_JSON.exists():
        return {}, meta

    payload = json.loads(SALES_SOURCE_JSON.read_text(encoding="utf-8"))
    shipper_rows = payload.get("shipper", [])
    meta["sourceRows"] = len(shipper_rows)

    weighted: dict[str, Counter] = {}
    for row in shipper_rows:
        origin_port = key_code(row.get("ori_port") or row.get("POR_PLC_CD"))
        booking_shipper = key_code(row.get("BKG_SHPR_CST_NO"))
        salesperson = normalize_salesperson(row.get("Salesman_POR"))
        if not origin_port or not booking_shipper or not salesperson:
            continue

        key = f"{origin_port}|{booking_shipper}"
        weight = number(row.get("fst")) or number(row.get("norm_lst")) or 1
        weighted.setdefault(key, Counter())[salesperson] += weight

    lookup: dict[str, str] = {}
    duplicate_keys = 0
    for key, counter in weighted.items():
        if len(counter) > 1:
            duplicate_keys += 1
        salesperson = sorted(
            counter.items(),
            key=lambda item: (-item[1], -len(item[0].split(",")), item[0]),
        )[0][0]
        lookup[key] = salesperson

    meta["mappingKeys"] = len(lookup)
    meta["duplicateKeys"] = duplicate_keys
    return lookup, meta


def efc_destination_rule(dest_country: str, dest_port: str) -> tuple[str, int, int] | None:
    if dest_country == "JP":
        return ("JAPAN", 60, 120)
    if dest_country in {"CN", "HK", "TW"}:
        return ("CHINA/HONG KONG/TAIWAN", 40, 80)
    if dest_country in SOUTH_EAST_ASIA:
        return ("SOUTH EAST ASIA", 40, 80)
    if dest_country in ISC:
        return ("INDIA/PAKISTAN (ISC)", 160, 320)
    if dest_country in RED_SEA_COUNTRIES or dest_port in RED_SEA_PORTS:
        return ("RED SEA", 200, 400)
    if dest_country in MIDDLE_EAST:
        return ("MIDDLE EAST", 160, 320)
    if dest_country in AFRICA:
        return ("AFRICA", 200, 400)
    if dest_country == "MX":
        return ("MEXICO", 200, 400)
    return None


def detect_source_encoding(path: Path) -> str:
    required_headers = {"실적년", "실적월", "실적년주차"}
    first_valid_encoding = SOURCE_ENCODINGS[0]

    for encoding in SOURCE_ENCODINGS:
        try:
            with path.open("r", encoding=encoding, newline="") as f:
                reader = csv.reader(f)
                headers = next(reader, [])
        except UnicodeDecodeError:
            continue

        if headers and first_valid_encoding == SOURCE_ENCODINGS[0]:
            first_valid_encoding = encoding
        if required_headers <= set(headers):
            return encoding

    return first_valid_encoding


def status_for(expected: float, actual: float) -> str:
    if expected <= 0 and actual > 0:
        return "수량확인"
    if expected <= 0:
        return "수량없음"
    if actual <= 0:
        return "미징수"
    if actual < expected * 0.95:
        return "부분징수"
    if actual > expected * 1.05:
        return "초과징수"
    return "정상"


def read_rows() -> tuple[list[dict], dict]:
    records: list[dict] = []
    skipped_summary = 0
    skipped_no_freight = 0
    raw_rows = 0
    salesperson_lookup, sales_meta = load_salesperson_lookup()
    sales_matched_rows = 0
    sales_unmatched_rows = 0
    sales_matched_keys: set[str] = set()
    sales_unmatched_keys: set[str] = set()

    with SOURCE_CSV.open("r", encoding=detect_source_encoding(SOURCE_CSV), newline="") as f:
        reader = csv.DictReader(f)
        for row_no, row in enumerate(reader, start=2):
            raw_rows += 1
            year = clean(row.get("실적년"))
            month = clean(row.get("실적월"))
            week_raw = clean(row.get("실적년주차"))

            # DynamicList exports a final total row with blank year/month and very large sums.
            if not year or not month:
                skipped_summary += 1
                continue

            of20 = number(row.get("20 o/f"))
            of40 = number(row.get("40 o/f"))
            if of20 == 0 and of40 == 0:
                skipped_no_freight += 1
                continue

            origin_country = clean(row.get("por국가"))
            origin_area = clean(row.get("porarea"))
            origin_port = clean(row.get("pol지역"))
            dest_country = clean(row.get("dly국가"))
            dest_area = clean(row.get("dlyarea"))
            dest_port = clean(row.get("dly지역"))

            qty20 = number(row.get("20갯수"))
            qty40 = number(row.get("40갯수"))
            teu = number(row.get("전체 teu"))
            lss_actual = number(row.get("20 lss")) + number(row.get("40 lss"))
            efc_actual = number(row.get("20 efc")) + number(row.get("40 efc"))

            program = "대상외"
            tariff_category = ""
            charge_basis = ""
            rate20 = 0
            rate40 = 0
            actual = 0.0
            expected = 0.0

            if origin_country == "CN" and dest_country == "JP":
                program = "LSS CN→JP"
                tariff_category = "CHINA TO JAPAN"
                charge_basis = "LSS increase"
                rate20 = 150
                rate40 = 300
                actual = lss_actual
                expected = qty20 * rate20 + qty40 * rate40
            elif origin_country not in EFC_EXCLUDED_ORIGINS:
                rule = efc_destination_rule(dest_country, dest_port)
                if rule:
                    tariff_category, rate20, rate40 = rule
                    program = "EFC non-CN"
                    charge_basis = "EFC tariff"
                    actual = efc_actual
                    expected = qty20 * rate20 + qty40 * rate40

            if program == "대상외":
                continue

            booking_shipper = clean(row.get("booking shipper"))
            handling_consignee = clean(row.get("handling consignee"))
            sales_key = f"{key_code(origin_port)}|{key_code(booking_shipper)}"
            salesperson = salesperson_lookup.get(sales_key, "") if booking_shipper and origin_port else ""
            if salesperson:
                sales_matched_rows += 1
                sales_matched_keys.add(sales_key)
            elif booking_shipper and origin_port:
                sales_unmatched_rows += 1
                sales_unmatched_keys.add(sales_key)
            status = status_for(expected, actual)
            gap = actual - expected

            records.append(
                {
                    "row": row_no,
                    "year": year,
                    "month": month.zfill(2),
                    "yearMonth": f"{year}-{month.zfill(2)}",
                    "week": str(int(number(week_raw))) if week_raw else "",
                    "originCountry": origin_country,
                    "originArea": origin_area,
                    "originPort": origin_port,
                    "destinationCountry": dest_country,
                    "destinationArea": dest_area,
                    "destinationPort": dest_port,
                    "bookingShipper": booking_shipper,
                    "handlingConsignee": handling_consignee,
                    "salesperson": salesperson,
                    "bl": clean(row.get("bl번호")),
                    "pc": clean(row.get("p/c")),
                    "qty20": whole(qty20),
                    "qty40": whole(qty40),
                    "teu": whole(teu),
                    "route": clean(row.get("route")),
                    "vessel": clean(row.get("vessel")),
                    "voyage": clean(row.get("voyage no")),
                    "cargoMode": clean(row.get("cgo mode")),
                    "program": program,
                    "tariffCategory": tariff_category,
                    "chargeBasis": charge_basis,
                    "rate20": rate20,
                    "rate40": rate40,
                    "expected": round(expected, 2),
                    "actual": round(actual, 2),
                    "gap": round(gap, 2),
                    "status": status,
                }
            )

    meta = {
        "source": SOURCE_CSV.name,
        "rawRows": raw_rows,
        "targetRows": len(records),
        "skippedSummaryRows": skipped_summary,
        "skippedNoFreightRows": skipped_no_freight,
        "programCounts": dict(Counter(r["program"] for r in records)),
        "statusCounts": dict(Counter(r["status"] for r in records)),
        "salesMapping": {
            **sales_meta,
            "matchedRows": sales_matched_rows,
            "unmatchedRows": sales_unmatched_rows,
            "matchedKeys": len(sales_matched_keys),
            "unmatchedKeys": len(sales_unmatched_keys),
        },
    }
    return records, meta


HTML = r"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>EFC/LSS Collection Dashboard</title>
  <style>
    :root {
      --bg: #f6f7f4;
      --panel: #ffffff;
      --ink: #1d2521;
      --muted: #69756f;
      --line: #d9dfda;
      --line-strong: #b8c4bc;
      --green: #1f7a5b;
      --green-soft: #dceee7;
      --blue: #276fbf;
      --blue-soft: #deebf8;
      --yellow: #f2b84b;
      --yellow-soft: #fff2cf;
      --red: #c84747;
      --red-soft: #f7dddd;
      --violet: #6a5acd;
      --shadow: 0 12px 30px rgba(29, 37, 33, 0.08);
      --radius: 8px;
    }

    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      font-size: 14px;
      letter-spacing: 0;
    }

    .login-wrap {
      display: flex;
      min-height: 100vh;
      align-items: center;
      justify-content: center;
      background: var(--bg);
      padding: 24px;
    }

    .login-box {
      width: min(420px, 100%);
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      padding: 32px;
      text-align: center;
    }

    .login-box h2 {
      margin: 0;
      color: var(--green);
      font-size: 22px;
      font-weight: 820;
    }

    .login-box p {
      margin: 8px 0 22px;
      color: var(--muted);
      line-height: 1.45;
    }

    .login-button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 10px;
      min-height: 42px;
      border: 0;
      border-radius: 6px;
      background: var(--green);
      color: #fff;
      padding: 0 18px;
      font: inherit;
      font-weight: 760;
      cursor: pointer;
    }

    .login-button:hover {
      background: #186448;
    }

    .login-error {
      display: none;
      margin: 14px 0 0;
      color: var(--red);
      font-size: 13px;
      font-weight: 720;
    }

    .login-help {
      margin-top: 14px;
      font-size: 13px;
    }

    .login-help a {
      color: var(--green);
      font-weight: 760;
      text-decoration: none;
    }

    .login-help a:hover {
      text-decoration: underline;
    }

    .lang-btn {
      cursor: pointer;
      font: inherit;
      font-size: 12px;
      font-weight: 760;
    }

    .login-help .lang-btn {
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fff;
      color: var(--green);
      min-height: 28px;
      padding: 0 12px;
    }

    .app-shell {
      display: none;
    }

    .app-message {
      margin-bottom: 14px;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      padding: 14px;
      color: var(--muted);
      font-weight: 760;
    }

    .app-message.error {
      color: var(--red);
      border-color: #e6b5b5;
      background: #fff8f8;
    }

    header {
      background: #17382c;
      color: #f8fbf8;
      padding: 22px 28px 18px;
      border-bottom: 4px solid #d4a12d;
    }

    .header-top {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
    }

    header h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 780;
      letter-spacing: 0;
    }

    .header-actions {
      display: flex;
      align-items: center;
      gap: 8px;
      flex-shrink: 0;
    }

    .header-link {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 30px;
      border: 1px solid rgba(248, 251, 248, 0.38);
      border-radius: 6px;
      padding: 0 12px;
      background: rgba(248, 251, 248, 0.12);
      color: #f8fbf8;
      text-decoration: none;
      font-size: 12px;
      font-weight: 760;
      white-space: nowrap;
    }

    .header-link:hover {
      background: rgba(248, 251, 248, 0.2);
    }

    .header-user {
      display: inline-flex;
      align-items: center;
      min-height: 30px;
      border: 0;
      background: transparent;
      color: rgba(248, 251, 248, 0.78);
      font: inherit;
      font-size: 12px;
      font-weight: 720;
      cursor: pointer;
      white-space: nowrap;
    }

    .header-user:hover {
      color: #fff;
    }

    header .sub {
      margin-top: 6px;
      color: #cbd9d1;
      display: flex;
      gap: 16px;
      flex-wrap: wrap;
      font-size: 13px;
    }

    main {
      padding: 18px 20px 28px;
      max-width: 1640px;
      margin: 0 auto;
    }

    .toolbar, .kpis, .grid, .wide-grid {
      display: grid;
      gap: 12px;
    }

    .toolbar {
      grid-template-columns: repeat(12, minmax(0, 1fr));
      align-items: end;
      margin-bottom: 14px;
    }

    .control {
      min-width: 0;
    }

    .control.small { grid-column: span 1; }
    .control.medium { grid-column: span 2; }
    .control.large { grid-column: span 3; }

    label {
      display: block;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      margin: 0 0 5px;
    }

    select, input {
      width: 100%;
      height: 36px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: var(--panel);
      color: var(--ink);
      padding: 0 10px;
      outline: none;
      font: inherit;
    }

    input:focus, select:focus {
      border-color: var(--green);
      box-shadow: 0 0 0 3px rgba(31, 122, 91, 0.14);
    }

    .segmented {
      display: inline-grid;
      width: 100%;
      grid-auto-flow: column;
      grid-auto-columns: 1fr;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: var(--panel);
      overflow: hidden;
      min-height: 36px;
    }

    .segmented button {
      border: 0;
      border-right: 1px solid var(--line);
      background: transparent;
      color: var(--muted);
      font: inherit;
      font-size: 13px;
      font-weight: 760;
      cursor: pointer;
      padding: 0 8px;
      min-width: 0;
    }

    .segmented button:last-child { border-right: 0; }
    .segmented button.active {
      background: var(--green);
      color: #fff;
    }

    .kpis {
      grid-template-columns: repeat(6, minmax(0, 1fr));
      margin-bottom: 14px;
    }

    .panel, .metric {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
    }

    .metric {
      padding: 14px 14px 12px;
      min-width: 0;
    }

    .metric .label {
      color: var(--muted);
      font-size: 12px;
      font-weight: 760;
    }

    .metric .value {
      margin-top: 8px;
      font-size: 24px;
      line-height: 1.05;
      font-weight: 820;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .metric .delta {
      margin-top: 7px;
      color: var(--muted);
      font-size: 12px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .grid {
      grid-template-columns: minmax(0, 1.25fr) minmax(360px, 0.75fr);
      margin-bottom: 14px;
    }

    .wide-grid {
      grid-template-columns: minmax(0, 1fr);
    }

    .panel {
      min-width: 0;
      overflow: hidden;
    }

    .panel-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 13px 14px;
      border-bottom: 1px solid var(--line);
    }

    .panel-title {
      font-size: 15px;
      font-weight: 820;
      min-width: 0;
    }

    .breadcrumb {
      display: flex;
      align-items: center;
      gap: 7px;
      flex-wrap: wrap;
      color: var(--muted);
      font-size: 12px;
      font-weight: 720;
    }

    .breadcrumb button {
      border: 1px solid var(--line);
      background: #fff;
      color: var(--ink);
      border-radius: 6px;
      padding: 5px 8px;
      cursor: pointer;
      font-size: 12px;
      font-weight: 760;
    }

    .chart {
      padding: 12px 14px 16px;
      min-height: 333px;
    }

    .bar-row {
      display: grid;
      grid-template-columns: minmax(86px, 160px) minmax(0, 1fr) 94px;
      gap: 10px;
      align-items: center;
      margin: 8px 0;
    }

    .bar-label {
      font-size: 12px;
      font-weight: 760;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .bar-track {
      position: relative;
      height: 22px;
      background: #edf1ee;
      border-radius: 4px;
      overflow: hidden;
    }

    .bar-fill {
      position: absolute;
      inset: 0 auto 0 0;
      background: linear-gradient(90deg, var(--red), var(--yellow));
      border-radius: 4px;
    }

    .bar-actual {
      position: absolute;
      inset: 4px auto 4px 0;
      background: var(--green);
      border-radius: 3px;
      opacity: 0.9;
    }

    .bar-value {
      color: var(--muted);
      text-align: right;
      font-size: 12px;
      font-variant-numeric: tabular-nums;
    }

    .sales-status {
      padding: 10px 12px 14px;
      min-height: 333px;
      overflow: auto;
    }

    .sales-status-summary {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 8px;
      margin-bottom: 10px;
    }

    .sales-status-summary div {
      border: 1px solid var(--line);
      border-radius: 7px;
      padding: 8px;
      background: #fbfcfb;
      min-width: 0;
    }

    .sales-status-summary span {
      display: block;
      color: var(--muted);
      font-size: 10px;
      font-weight: 800;
      text-transform: uppercase;
    }

    .sales-status-summary strong {
      display: block;
      margin-top: 3px;
      font-size: 15px;
      font-variant-numeric: tabular-nums;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .sales-status-table {
      min-width: 640px;
      font-size: 12px;
    }

    .sales-status-table th,
    .sales-status-table td {
      padding: 7px 6px;
      height: 32px;
      border-bottom: 1px solid var(--line);
      vertical-align: middle;
    }

    .sales-status-table th {
      font-size: 10px;
    }

    .sales-status-table .sales-name {
      max-width: 86px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      font-weight: 800;
    }

    .sales-gap-cell {
      min-width: 86px;
    }

    .sales-gap-text {
      display: flex;
      justify-content: flex-end;
      font-variant-numeric: tabular-nums;
      font-weight: 760;
      margin-bottom: 3px;
    }

    .sales-gap-bar {
      height: 5px;
      border-radius: 999px;
      background: #edf1ee;
      overflow: hidden;
    }

    .sales-gap-bar span {
      display: block;
      height: 100%;
      width: 0;
      border-radius: inherit;
      background: linear-gradient(90deg, var(--red), var(--yellow));
    }

    .table-wrap {
      overflow: auto;
      max-height: 560px;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      min-width: 940px;
    }

    th, td {
      padding: 10px 10px;
      border-bottom: 1px solid var(--line);
      text-align: left;
      vertical-align: middle;
      white-space: nowrap;
    }

    th {
      position: sticky;
      top: 0;
      z-index: 1;
      background: #f8faf8;
      color: #53615a;
      font-size: 12px;
      font-weight: 820;
      cursor: pointer;
      user-select: none;
    }

    td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
    tr.data-row { cursor: pointer; }
    tr.data-row:hover { background: #f4faf7; }

    .pill {
      display: inline-flex;
      align-items: center;
      height: 24px;
      border-radius: 999px;
      padding: 0 8px;
      font-size: 12px;
      font-weight: 780;
      border: 1px solid transparent;
    }

    .pill.ok { background: var(--green-soft); color: var(--green); }
    .pill.missing { background: var(--red-soft); color: var(--red); }
    .pill.partial { background: var(--yellow-soft); color: #916113; }
    .pill.over { background: var(--blue-soft); color: var(--blue); }
    .pill.check { background: #eee9ff; color: var(--violet); }
    .pill.empty-qty { background: #ecefed; color: var(--muted); }

    .rules {
      padding: 12px 14px 14px;
      display: grid;
      gap: 10px;
    }

    .rule-table {
      min-width: 0;
      font-size: 12px;
    }

    .rule-table th, .rule-table td {
      padding: 7px 8px;
    }

    .note {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.45;
      padding-top: 3px;
    }

    .empty {
      padding: 32px 14px;
      color: var(--muted);
      text-align: center;
      font-weight: 740;
    }

    @media (max-width: 1180px) {
      .toolbar { grid-template-columns: repeat(6, minmax(0, 1fr)); }
      .control.small, .control.medium, .control.large { grid-column: span 2; }
      .kpis { grid-template-columns: repeat(3, minmax(0, 1fr)); }
      .grid { grid-template-columns: 1fr; }
    }

    @media (max-width: 720px) {
      header { padding: 18px 16px 14px; }
      .header-top { align-items: flex-start; flex-direction: column; }
      main { padding: 14px 12px 22px; }
      .toolbar { grid-template-columns: 1fr; }
      .control.small, .control.medium, .control.large { grid-column: span 1; }
      .kpis { grid-template-columns: 1fr; }
      .metric .value { font-size: 22px; }
      .bar-row { grid-template-columns: 92px minmax(0, 1fr) 74px; }
      .panel-head { align-items: flex-start; flex-direction: column; }
    }
  </style>
</head>
<body>
  <div class="login-wrap" id="login">
    <div class="login-box">
      <h2>EFC/LSS Collection Dashboard</h2>
      <p id="loginIntro">회사 Google 계정으로 로그인하면 데이터를 볼 수 있습니다.</p>
      <button class="login-button" onclick="doLogin()">
        <svg width="18" height="18" viewBox="0 0 48 48" aria-hidden="true">
          <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/>
          <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/>
          <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/>
          <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/>
        </svg>
        <span id="loginButtonText">Google 계정으로 로그인</span>
      </button>
      <div class="login-help"><a href="guide.html" id="loginGuideLink">사용가이드 보기</a></div>
      <div class="login-help"><button class="lang-btn" onclick="toggleLang()" type="button">EN</button></div>
      <p class="login-error" id="loginErr"></p>
    </div>
  </div>

  <div id="app" class="app-shell">
    <header>
      <div class="header-top">
        <h1 id="dashboardTitle">EFC/LSS Tariff Collection Monitor</h1>
        <div class="header-actions">
          <button class="header-link lang-btn" onclick="toggleLang()" type="button">EN</button>
          <a class="header-link" href="guide.html" id="guideLink">Guide</a>
          <button class="header-user" id="userInfo" onclick="logout()"></button>
        </div>
      </div>
      <div class="sub">
        <span id="sourceMeta"></span>
        <span id="lssBasisText">CN→JP LSS USD 150/300</span>
        <span id="efcBasisText">EFC DRY tariff basis</span>
      </div>
    </header>

    <main>
      <div class="app-message" id="loading">데이터 로딩 중...</div>

      <section class="toolbar" aria-label="filters">
      <div class="control large">
        <label>Charge</label>
        <div class="segmented" data-segment="program">
          <button data-value="ALL" class="active">전체</button>
          <button data-value="EFC non-CN">EFC</button>
          <button data-value="LSS CN→JP">LSS</button>
        </div>
      </div>
      <div class="control medium">
        <label>Layer</label>
        <div class="segmented" data-segment="level">
          <button data-value="origin" class="active">선적지</button>
          <button data-value="lane">Lane</button>
          <button data-value="customer">고객</button>
          <button data-value="salesperson">영업사원</button>
        </div>
      </div>
      <div class="control medium">
        <label>Origin</label>
        <div class="segmented" data-segment="originBasis">
          <button data-value="originPort">POL</button>
          <button data-value="originCountry" class="active">국가</button>
        </div>
      </div>
      <div class="control medium">
        <label>Customer</label>
        <div class="segmented" data-segment="customerBasis">
          <button data-value="bookingShipper" class="active">Shipper</button>
          <button data-value="handlingConsignee">CNEE</button>
        </div>
      </div>
      <div class="control small">
        <label>Month</label>
        <select id="monthFilter"></select>
      </div>
      <div class="control small">
        <label>Week</label>
        <select id="weekFilter"></select>
      </div>
      <div class="control small">
        <label>P/C</label>
        <select id="pcFilter"></select>
      </div>
      <div class="control medium">
        <label>Status</label>
        <select id="statusFilter"></select>
      </div>
      <div class="control medium">
        <label>Salesperson</label>
        <select id="salespersonFilter"></select>
      </div>
      <div class="control medium">
        <label>선적지</label>
        <select id="originFilter"></select>
      </div>
      <div class="control medium">
        <label>도착지</label>
        <select id="destinationFilter"></select>
      </div>
      <div class="control large">
        <label>Search</label>
        <input id="searchFilter" type="search" placeholder="BL / 고객 / Port / Route">
      </div>
      </section>

      <section class="kpis">
      <div class="metric">
        <div class="label">징수율</div>
        <div class="value" id="kpiRate">-</div>
        <div class="delta" id="kpiRateDelta">-</div>
      </div>
      <div class="metric">
        <div class="label">Tariff 기대액</div>
        <div class="value" id="kpiExpected">-</div>
        <div class="delta" id="kpiExpectedDelta">-</div>
      </div>
      <div class="metric">
        <div class="label">실제 징수액</div>
        <div class="value" id="kpiActual">-</div>
        <div class="delta" id="kpiActualDelta">-</div>
      </div>
      <div class="metric">
        <div class="label">Gap</div>
        <div class="value" id="kpiGap">-</div>
        <div class="delta" id="kpiGapDelta">-</div>
      </div>
      <div class="metric">
        <div class="label">대상 BL</div>
        <div class="value" id="kpiBl">-</div>
        <div class="delta" id="kpiBlDelta">-</div>
      </div>
      <div class="metric">
        <div class="label">TEU</div>
        <div class="value" id="kpiTeu">-</div>
        <div class="delta" id="kpiTeuDelta">-</div>
      </div>
      </section>

      <section class="grid">
      <div class="panel">
        <div class="panel-head">
          <div>
            <div class="panel-title" id="mainTitle">선적지별 징수율</div>
            <div class="breadcrumb" id="breadcrumb"></div>
          </div>
          <div class="segmented" data-segment="sortMetric" style="width: 290px;">
            <button data-value="gap" class="active">Gap</button>
            <button data-value="expected">Tariff</button>
            <button data-value="rate">징수율</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="mainTable"></table>
        </div>
      </div>

      <div class="panel">
        <div class="panel-head">
          <div class="panel-title" id="salesStatusTitle">영업사원별 현황</div>
        </div>
        <div class="sales-status" id="salesStatus"></div>
      </div>
      </section>

      <section class="grid">
      <div class="panel">
        <div class="panel-head">
          <div class="panel-title">BL Exception</div>
        </div>
        <div class="table-wrap">
          <table id="exceptionTable"></table>
        </div>
      </div>

      <div class="panel">
        <div class="panel-head">
          <div class="panel-title" id="tariffBasisTitle">Tariff Basis</div>
        </div>
        <div class="rules">
          <table class="rule-table">
            <thead>
              <tr><th id="tariffCategoryHead">구분</th><th class="num">20 DRY</th><th class="num">40 DRY</th></tr>
            </thead>
            <tbody>
              <tr><td>LSS CN→JP</td><td class="num">150</td><td class="num">300</td></tr>
              <tr><td>EFC JAPAN</td><td class="num">60</td><td class="num">120</td></tr>
              <tr><td>EFC CHINA/HK/TAIWAN</td><td class="num">40</td><td class="num">80</td></tr>
              <tr><td>EFC SOUTH EAST ASIA</td><td class="num">40</td><td class="num">80</td></tr>
              <tr><td>EFC INDIA/PAKISTAN (ISC)</td><td class="num">160</td><td class="num">320</td></tr>
              <tr><td>EFC MIDDLE EAST</td><td class="num">160</td><td class="num">320</td></tr>
              <tr><td>EFC RED SEA / AFRICA / MEXICO</td><td class="num">200</td><td class="num">400</td></tr>
            </tbody>
          </table>
          <div class="note" id="tariffNote">
            LSS effective date: 2026-03-23. EFC effective date: 2026-03-28, Vietnam origin: 2026-04-01.
            The source file has year/month/week but no ETD POL date, so row-level effective-date cutoff is not applied.
            RF/RH 50% uplift is not applied because the source file has no refrigerated cargo flag.
          </div>
        </div>
      </div>
      </section>
    </main>
  </div>

  <script>
    const DATA_URL = "data.json";
    const GOOGLE_CLIENT_ID = "409330651463-giie223egsskdq10etn642gjtron1hq5.apps.googleusercontent.com";
    const REDIRECT_PATH = "/EFC_LSS_monitoring/";
    const ALLOWED_DOMAINS = ["ekmtc.com"];
    const SCOPES = "email profile openid";
    let rows = [];
    let meta = {};
    let gToken = null;
    let gUser = null;
    let appInitialized = false;
    let lang = localStorage.getItem("efcLssLang") || localStorage.getItem("lang") || "ko";

    const I18N = {
      ko: {
        langButton: "EN",
        dashboardTitle: "EFC/LSS Tariff Collection Monitor",
        loginIntro: "회사 Google 계정으로 로그인하면 데이터를 볼 수 있습니다.",
        loginButton: "Google 계정으로 로그인",
        loginGuide: "사용가이드 보기",
        guide: "Guide",
        logout: "로그아웃",
        loading: "데이터 로딩 중...",
        loadFailed: "데이터 로드 실패",
        fileLoginError: "Google 로그인은 GitHub Pages URL에서 실행해야 합니다.",
        userInfoFailed: "사용자 정보 확인 실패",
        domainDenied: domain => `${domain || "Unknown"} 도메인은 접근할 수 없습니다.`,
        lssBasis: "CN→JP LSS USD 150/300",
        efcBasis: "EFC DRY tariff basis",
        salesStatusTitle: "영업사원별 현황",
        sourceMeta: meta => {
          const sales = meta.salesMapping && meta.salesMapping.found
            ? ` · 영업사원 ${num(meta.salesMapping.matchedRows)} rows 매핑`
            : " · 영업사원 매핑 없음";
          return `${meta.source || "data.json"} · 대상 ${num(meta.targetRows)} rows · no-O/F 제외 ${num(meta.skippedNoFreightRows)} rows${sales}`;
        },
        filters: {
          charge: "Charge", layer: "Layer", origin: "Origin", customer: "Customer", month: "Month",
          week: "Week", pc: "P/C", status: "Status", salesperson: "영업사원", originSelect: "선적지", destination: "도착지", search: "Search",
        },
        buttons: {
          all: "전체", origin: "선적지", lane: "Lane", customer: "고객", salesperson: "영업사원", country: "국가",
          shipper: "Shipper", cnee: "CNEE", gap: "Gap", tariff: "Tariff", rate: "징수율",
        },
        searchPlaceholder: "BL / 고객 / 영업사원 / Port / Route",
        kpis: {
          rate: "징수율", expected: "Tariff 기대액", actual: "실제 징수액", gap: "Gap", bl: "대상 BL", teu: "TEU",
          underCollected: count => `${num(count)} under-collected rows`,
          qty: (qty20, qty40) => `20' ${num(qty20)} / 40' ${num(qty40)}`,
          overCheck: (over, check) => `${num(over)} over / ${num(check)} qty-check`,
          shortfall: "shortfall",
          overCollected: "over-collected",
          rows: count => `${num(count)} rows`,
        },
        labels: {
          origin: "선적지", lane: "선적지-도착지", customer: "고객", salesperson: "영업사원",
          expected: "Tariff", actual: "징수", gap: "Gap", rate: "징수율",
        },
        table: {
          titleSuffix: "별 징수율", noData: "데이터 없음", noException: "예외 없음",
          bl: "BL", teu: "TEU", issue: "미/부분", category: "구분", status: "Status",
          bookingShipper: "Booking Shipper", salesperson: "영업사원", charge: "Charge", pol: "POL", pod: "POD", tariffCat: "Tariff Cat.",
          salespersonStatus: "영업사원", shipperCount: "업체수", missingShipperCount: "미징수 업체", partialShipperCount: "부분 업체",
          mappedShippers: "업체수", owners: "담당자",
        },
        tariffBasis: "Tariff Basis",
        tariffCategory: "구분",
        tariffNote: "LSS effective date: 2026-03-23. EFC effective date: 2026-03-28, Vietnam origin: 2026-04-01. The source file has year/month/week but no ETD POL date, so row-level effective-date cutoff is not applied. RF/RH 50% uplift is not applied because the source file has no refrigerated cargo flag.",
        statusMap: {
          "정상": "정상",
          "미징수": "미징수",
          "부분징수": "부분징수",
          "초과징수": "초과징수",
          "수량확인": "수량확인",
          "수량없음": "수량없음",
        },
        unassigned: "미지정",
      },
      en: {
        langButton: "KO",
        dashboardTitle: "EFC/LSS Tariff Collection Monitor",
        loginIntro: "Sign in with your company Google account to view the dashboard.",
        loginButton: "Sign in with Google",
        loginGuide: "View User Guide",
        guide: "Guide",
        logout: "Logout",
        loading: "Loading data...",
        loadFailed: "Data load failed",
        fileLoginError: "Google login must be used from the GitHub Pages URL.",
        userInfoFailed: "Failed to verify user",
        domainDenied: domain => `${domain || "Unknown"} domain is not allowed.`,
        lssBasis: "CN→JP LSS USD 150/300",
        efcBasis: "EFC DRY tariff basis",
        salesStatusTitle: "Salesperson Status",
        sourceMeta: meta => {
          const sales = meta.salesMapping && meta.salesMapping.found
            ? ` · ${num(meta.salesMapping.matchedRows)} rows sales-mapped`
            : " · no salesperson mapping";
          return `${meta.source || "data.json"} · ${num(meta.targetRows)} target rows · ${num(meta.skippedNoFreightRows)} no-O/F rows skipped${sales}`;
        },
        filters: {
          charge: "Charge", layer: "Layer", origin: "Origin", customer: "Customer", month: "Month",
          week: "Week", pc: "P/C", status: "Status", salesperson: "Salesperson", originSelect: "Origin", destination: "Destination", search: "Search",
        },
        buttons: {
          all: "All", origin: "Origin", lane: "Lane", customer: "Customer", salesperson: "Sales", country: "Country",
          shipper: "Shipper", cnee: "CNEE", gap: "Gap", tariff: "Tariff", rate: "Rate",
        },
        searchPlaceholder: "BL / customer / salesperson / port / route",
        kpis: {
          rate: "Collection Rate", expected: "Tariff Expected", actual: "Actual Collection", gap: "Gap", bl: "Target BL", teu: "TEU",
          underCollected: count => `${num(count)} under-collected rows`,
          qty: (qty20, qty40) => `20' ${num(qty20)} / 40' ${num(qty40)}`,
          overCheck: (over, check) => `${num(over)} over / ${num(check)} qty-check`,
          shortfall: "shortfall",
          overCollected: "over-collected",
          rows: count => `${num(count)} rows`,
        },
        labels: {
          origin: "Origin", lane: "Lane", customer: "Customer", salesperson: "Salesperson",
          expected: "Tariff", actual: "Actual", gap: "Gap", rate: "Rate",
        },
        table: {
          titleSuffix: " Collection Rate", noData: "No data", noException: "No exception",
          bl: "BL", teu: "TEU", issue: "Missing/Partial", category: "Category", status: "Status",
          bookingShipper: "Booking Shipper", salesperson: "Salesperson", charge: "Charge", pol: "POL", pod: "POD", tariffCat: "Tariff Cat.",
          salespersonStatus: "Salesperson", shipperCount: "Shippers", missingShipperCount: "Missing Shippers", partialShipperCount: "Partial Shippers",
          mappedShippers: "Shippers", owners: "Owners",
        },
        tariffBasis: "Tariff Basis",
        tariffCategory: "Category",
        tariffNote: "LSS effective date: 2026-03-23. EFC effective date: 2026-03-28, Vietnam origin: 2026-04-01. The source file has year/month/week but no ETD POL date, so row-level effective-date cutoff is not applied. RF/RH 50% uplift is not applied because the source file has no refrigerated cargo flag.",
        statusMap: {
          "정상": "Normal",
          "미징수": "Missing",
          "부분징수": "Partial",
          "초과징수": "Over",
          "수량확인": "Qty Check",
          "수량없음": "No Qty",
        },
        unassigned: "Unassigned",
      },
    };

    const state = {
      program: "ALL",
      level: "origin",
      originBasis: "originCountry",
      customerBasis: "bookingShipper",
      sortMetric: "gap",
      month: "ALL",
      week: "ALL",
      pc: "ALL",
      status: "ALL",
      salesperson: "ALL",
      origin: "ALL",
      destination: "ALL",
      search: "",
      selectedOrigin: "",
      selectedDestination: "",
      tableSort: { key: "gap", direction: "asc" },
    };

    function t() {
      return I18N[lang] || I18N.ko;
    }

    function usd(value) {
      const rounded = Math.round(value || 0);
      const sign = rounded < 0 ? "-" : "";
      return sign + "$" + Math.abs(rounded).toLocaleString("en-US");
    }

    function signedUsd(value) {
      const rounded = Math.round(value || 0);
      const sign = rounded > 0 ? "+" : rounded < 0 ? "-" : "";
      return sign + "$" + Math.abs(rounded).toLocaleString("en-US");
    }

    function num(value) {
      return Math.round(value || 0).toLocaleString("en-US");
    }

    function pct(value) {
      if (!Number.isFinite(value)) return "-";
      return (value * 100).toFixed(1) + "%";
    }

    function statusText(status) {
      return t().statusMap[status] || status || "-";
    }

    function safe(value, fallback = "-") {
      return value && String(value).trim() ? value : fallback;
    }

    function statusClass(status) {
      return {
        "정상": "ok",
        "미징수": "missing",
        "부분징수": "partial",
        "초과징수": "over",
        "수량확인": "check",
        "수량없음": "empty-qty",
      }[status] || "partial";
    }

    function originValue(row) {
      return state.originBasis === "originCountry" ? row.originCountry : row.originPort;
    }

    function originLabel(row) {
      if (state.originBasis === "originCountry") return safe(row.originCountry);
      return `${safe(row.originPort)} (${safe(row.originCountry)})`;
    }

    function destinationLabel(row) {
      return `${safe(row.destinationPort)} (${safe(row.destinationCountry)})`;
    }

    function customerValue(row) {
      const value = row[state.customerBasis];
      return safe(value, t().unassigned);
    }

    function salespersonValue(row) {
      return safe(row.salesperson, t().unassigned);
    }

    function matchesSearch(row, term) {
      if (!term) return true;
      const text = [
        row.bl, row.bookingShipper, row.handlingConsignee, row.originCountry, row.originPort,
        row.salesperson, row.destinationCountry, row.destinationPort, row.route, row.vessel, row.voyage,
        row.tariffCategory, row.status,
      ].join(" ").toLowerCase();
      return text.includes(term);
    }

    function filteredRows() {
      const term = state.search.trim().toLowerCase();
      return rows.filter(row => {
        if (state.program !== "ALL" && row.program !== state.program) return false;
        if (state.month !== "ALL" && row.yearMonth !== state.month) return false;
        if (state.week !== "ALL" && row.week !== state.week) return false;
        if (state.pc !== "ALL" && row.pc !== state.pc) return false;
        if (state.status !== "ALL" && row.status !== state.status) return false;
        if (state.salesperson !== "ALL" && salespersonValue(row) !== state.salesperson) return false;
        if (state.origin !== "ALL" && originValue(row) !== state.origin) return false;
        if (state.destination !== "ALL" && row.destinationPort !== state.destination) return false;
        if (state.selectedOrigin && originValue(row) !== state.selectedOrigin) return false;
        if (state.selectedDestination && row.destinationPort !== state.selectedDestination) return false;
        return matchesSearch(row, term);
      });
    }

    function aggregate(sourceRows, groupers) {
      const map = new Map();
      for (const row of sourceRows) {
        const parts = groupers.map(g => g.value(row));
        const key = parts.join("\u001f");
        if (!map.has(key)) {
          const labels = groupers.map(g => g.label(row));
          map.set(key, {
            key,
            parts,
            labels,
            rows: 0,
            blSet: new Set(),
            shipperSet: new Set(),
            issueShipperSet: new Set(),
            missingShipperSet: new Set(),
            partialShipperSet: new Set(),
            qty20: 0,
            qty40: 0,
            teu: 0,
            expected: 0,
            actual: 0,
            gap: 0,
            issueGap: 0,
            missing: 0,
            partial: 0,
            over: 0,
            check: 0,
            programs: new Set(),
            categories: new Set(),
            salespeople: new Set(),
          });
        }
        const item = map.get(key);
        item.rows += 1;
        item.blSet.add(row.bl);
        if (row.bookingShipper) item.shipperSet.add(row.bookingShipper);
        item.qty20 += Number(row.qty20 || 0);
        item.qty40 += Number(row.qty40 || 0);
        item.teu += Number(row.teu || 0);
        item.expected += Number(row.expected || 0);
        item.actual += Number(row.actual || 0);
        item.gap += Number(row.gap || 0);
        if (["미징수", "부분징수"].includes(row.status)) {
          item.issueGap += Number(row.gap || 0);
          if (row.bookingShipper) item.issueShipperSet.add(row.bookingShipper);
        }
        if (row.status === "미징수") {
          item.missing += 1;
          if (row.bookingShipper) item.missingShipperSet.add(row.bookingShipper);
        }
        if (row.status === "부분징수") {
          item.partial += 1;
          if (row.bookingShipper) item.partialShipperSet.add(row.bookingShipper);
        }
        if (row.status === "초과징수") item.over += 1;
        if (row.status === "수량확인") item.check += 1;
        item.programs.add(row.program);
        item.categories.add(row.tariffCategory);
        if (row.salesperson) item.salespeople.add(row.salesperson);
      }
      return Array.from(map.values()).map(item => ({
        ...item,
        bl: item.blSet.size,
        shippers: item.shipperSet.size,
        issueShippers: item.issueShipperSet.size,
        missingShippers: item.missingShipperSet.size,
        partialShippers: item.partialShipperSet.size,
        rate: item.expected > 0 ? item.actual / item.expected : NaN,
        programText: Array.from(item.programs).join(", "),
        categoryText: Array.from(item.categories).slice(0, 3).join(", "),
        salesText: Array.from(item.salespeople).sort().slice(0, 3).join(", ") + (item.salespeople.size > 3 ? "..." : ""),
        sortLabel: item.labels.join(" > "),
      }));
    }

    function currentGroupers() {
      const origin = {
        value: row => originValue(row),
        label: row => originLabel(row),
      };
      const destination = {
        value: row => row.destinationPort,
        label: row => destinationLabel(row),
      };
      const customer = {
        value: row => customerValue(row),
        label: row => customerValue(row),
      };
      const salesperson = {
        value: row => salespersonValue(row),
        label: row => salespersonValue(row),
      };
      if (state.level === "origin") return [origin];
      if (state.level === "lane") return [origin, destination];
      if (state.level === "salesperson") return [origin, destination, salesperson];
      return [origin, destination, customer];
    }

    function totals(sourceRows) {
      const bl = new Set();
      const total = sourceRows.reduce((acc, row) => {
        bl.add(row.bl);
        acc.expected += Number(row.expected || 0);
        acc.actual += Number(row.actual || 0);
        acc.gap += Number(row.gap || 0);
        acc.teu += Number(row.teu || 0);
        acc.qty20 += Number(row.qty20 || 0);
        acc.qty40 += Number(row.qty40 || 0);
        if (row.status === "미징수") acc.missing += 1;
        if (row.status === "부분징수") acc.partial += 1;
        if (row.status === "초과징수") acc.over += 1;
        if (row.status === "수량확인") acc.check += 1;
        return acc;
      }, { expected: 0, actual: 0, gap: 0, teu: 0, qty20: 0, qty40: 0, missing: 0, partial: 0, over: 0, check: 0 });
      total.bl = bl.size;
      total.rate = total.expected > 0 ? total.actual / total.expected : NaN;
      return total;
    }

    function fillSelect(id, values, current, formatter = x => x) {
      const select = document.getElementById(id);
      const previous = current || select.value || "ALL";
      const unique = Array.from(new Set(values.filter(Boolean))).sort((a, b) => String(a).localeCompare(String(b)));
      select.innerHTML = `<option value="ALL">${escapeHtml(t().buttons.all)}</option>` + unique.map(value => {
        const selected = value === previous ? " selected" : "";
        return `<option value="${escapeHtml(value)}"${selected}>${escapeHtml(formatter(value))}</option>`;
      }).join("");
      if (previous !== "ALL" && !unique.includes(previous)) {
        select.value = "ALL";
        return "ALL";
      }
      select.value = previous;
      return previous;
    }

    function escapeHtml(value) {
      return String(value ?? "").replace(/[&<>"']/g, ch => ({
        "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
      }[ch]));
    }

    function setText(id, value) {
      const element = document.getElementById(id);
      if (element) element.textContent = value;
    }

    function toggleLang() {
      lang = lang === "ko" ? "en" : "ko";
      localStorage.setItem("efcLssLang", lang);
      localStorage.setItem("lang", lang);
      applyLang();
      if (appInitialized) {
        setupFilters();
        syncSegments();
        render();
      }
    }

    function applyLang() {
      const ui = t();
      document.documentElement.lang = lang;
      document.querySelectorAll(".lang-btn").forEach(button => {
        button.textContent = ui.langButton;
      });
      setText("loginIntro", ui.loginIntro);
      setText("loginButtonText", ui.loginButton);
      setText("loginGuideLink", ui.loginGuide);
      setText("guideLink", ui.guide);
      setText("dashboardTitle", ui.dashboardTitle);
      setText("lssBasisText", ui.lssBasis);
      setText("efcBasisText", ui.efcBasis);
      setText("salesStatusTitle", ui.salesStatusTitle);
      setText("tariffBasisTitle", ui.tariffBasis);
      setText("tariffCategoryHead", ui.tariffCategory);
      setText("tariffNote", ui.tariffNote);

      const filterLabels = document.querySelectorAll(".toolbar .control > label");
      [
        ui.filters.charge, ui.filters.layer, ui.filters.origin, ui.filters.customer,
        ui.filters.month, ui.filters.week, ui.filters.pc, ui.filters.status,
        ui.filters.salesperson, ui.filters.originSelect, ui.filters.destination, ui.filters.search,
      ].forEach((label, index) => {
        if (filterLabels[index]) filterLabels[index].textContent = label;
      });

      const buttonLabels = {
        program: { ALL: ui.buttons.all, "EFC non-CN": "EFC", "LSS CN→JP": "LSS" },
        level: { origin: ui.buttons.origin, lane: ui.buttons.lane, customer: ui.buttons.customer, salesperson: ui.buttons.salesperson },
        originBasis: { originPort: "POL", originCountry: ui.buttons.country },
        customerBasis: { bookingShipper: ui.buttons.shipper, handlingConsignee: ui.buttons.cnee },
        sortMetric: { gap: ui.buttons.gap, expected: ui.buttons.tariff, rate: ui.buttons.rate },
      };
      document.querySelectorAll("[data-segment]").forEach(group => {
        const segment = group.dataset.segment;
        group.querySelectorAll("button").forEach(button => {
          const label = buttonLabels[segment]?.[button.dataset.value];
          if (label) button.textContent = label;
        });
      });

      const search = document.getElementById("searchFilter");
      if (search) search.placeholder = ui.searchPlaceholder;
      const kpiLabels = document.querySelectorAll(".kpis .metric .label");
      [ui.kpis.rate, ui.kpis.expected, ui.kpis.actual, ui.kpis.gap, ui.kpis.bl, ui.kpis.teu]
        .forEach((label, index) => {
          if (kpiLabels[index]) kpiLabels[index].textContent = label;
        });
      const loading = document.getElementById("loading");
      if (loading && loading.style.display !== "none" && !loading.classList.contains("error")) {
        loading.textContent = ui.loading;
      }
      if (gUser) {
        const userName = gUser?.name || gUser?.email || "User";
        document.getElementById("userInfo").textContent = `${userName} | ${ui.logout}`;
      }
    }

    function setLoginError(message) {
      const element = document.getElementById("loginErr");
      element.textContent = message;
      element.style.display = message ? "block" : "none";
    }

    function authRedirectUri() {
      return location.origin + REDIRECT_PATH;
    }

    function doLogin() {
      if (location.protocol === "file:") {
        setLoginError(t().fileLoginError);
        return;
      }

      const authUrl = "https://accounts.google.com/o/oauth2/v2/auth" +
        "?client_id=" + encodeURIComponent(GOOGLE_CLIENT_ID) +
        "&redirect_uri=" + encodeURIComponent(authRedirectUri()) +
        "&response_type=token" +
        "&scope=" + encodeURIComponent(SCOPES) +
        "&include_granted_scopes=true" +
        "&prompt=select_account";
      location.href = authUrl;
    }

    function handleRedirect() {
      const hash = location.hash.substring(1);
      if (!hash) return false;

      const params = new URLSearchParams(hash);
      const token = params.get("access_token");
      if (!token) return false;

      history.replaceState(null, "", location.pathname + location.search);
      fetch("https://www.googleapis.com/oauth2/v2/userinfo", {
        headers: { Authorization: "Bearer " + token },
      })
        .then(response => {
          if (!response.ok) throw new Error(`HTTP ${response.status}`);
          return response.json();
        })
        .then(info => {
          const domain = (info.email || "").split("@")[1] || "";
          if (ALLOWED_DOMAINS.length && !ALLOWED_DOMAINS.includes(domain.toLowerCase())) {
            setLoginError(t().domainDenied(domain));
            return;
          }

          gToken = token;
          gUser = { email: info.email, name: info.name, picture: info.picture };
          sessionStorage.setItem("efcLssGoogleToken", token);
          sessionStorage.setItem("efcLssGoogleUser", JSON.stringify(gUser));
          showApp();
        })
        .catch(error => {
          setLoginError(`${t().userInfoFailed}: ${error.message}`);
        });
      return true;
    }

    function checkSession() {
      if (handleRedirect()) return;

      const token = sessionStorage.getItem("efcLssGoogleToken");
      const user = sessionStorage.getItem("efcLssGoogleUser");
      if (!token || !user) return;

      gToken = token;
      try {
        gUser = JSON.parse(user);
      } catch {
        sessionStorage.removeItem("efcLssGoogleToken");
        sessionStorage.removeItem("efcLssGoogleUser");
        return;
      }

      fetch("https://www.googleapis.com/oauth2/v1/tokeninfo?access_token=" + encodeURIComponent(token))
        .then(response => {
          if (response.ok) showApp();
          else logout(false);
        })
        .catch(() => logout(false));
    }

    function logout(reload = true) {
      sessionStorage.removeItem("efcLssGoogleToken");
      sessionStorage.removeItem("efcLssGoogleUser");
      gToken = null;
      gUser = null;
      if (reload) location.reload();
    }

    async function loadData() {
      const loading = document.getElementById("loading");
      loading.classList.remove("error");
      loading.style.display = "block";
      loading.textContent = t().loading;

      try {
        const response = await fetch(`${DATA_URL}?t=${Date.now()}`);
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const payload = await response.json();
        rows = payload.rows || [];
        meta = payload.meta || {};
        loading.style.display = "none";
      } catch (error) {
        loading.classList.add("error");
        loading.textContent = `${t().loadFailed}: ${error.message}`;
        throw error;
      }
    }

    function showApp() {
      setLoginError("");
      document.getElementById("login").style.display = "none";
      document.getElementById("app").style.display = "block";
      const userName = gUser?.name || gUser?.email || "User";
      document.getElementById("userInfo").textContent = `${userName} | ${t().logout}`;

      if (!appInitialized) {
        loadData()
          .then(() => {
            setupFilters();
            syncSegments();
            render();
            appInitialized = true;
          })
          .catch(() => {});
      }
    }

    function setupFilters() {
      fillSelect("monthFilter", rows.map(r => r.yearMonth), state.month);
      fillSelect("weekFilter", rows.map(r => r.week), state.week, v => `W${v}`);
      fillSelect("pcFilter", rows.map(r => r.pc), state.pc);
      fillSelect("statusFilter", ["정상", "부분징수", "미징수", "초과징수", "수량확인", "수량없음"], state.status, statusText);
      fillSelect("salespersonFilter", rows.map(salespersonValue), state.salesperson);
      updateOriginDestinationFilters();
    }

    function updateOriginDestinationFilters() {
      const base = rows.filter(row => state.program === "ALL" || row.program === state.program);
      state.origin = fillSelect("originFilter", base.map(originValue), state.origin);
      state.destination = fillSelect("destinationFilter", base.map(r => r.destinationPort), state.destination, value => {
        const found = base.find(r => r.destinationPort === value);
        return found ? `${value} (${found.destinationCountry})` : value;
      });
    }

    function renderKpis(sourceRows) {
      const total = totals(sourceRows);
      document.getElementById("kpiRate").textContent = pct(total.rate);
      document.getElementById("kpiRateDelta").textContent = t().kpis.underCollected(total.missing + total.partial);
      document.getElementById("kpiExpected").textContent = usd(total.expected);
      document.getElementById("kpiExpectedDelta").textContent = t().kpis.qty(total.qty20, total.qty40);
      document.getElementById("kpiActual").textContent = usd(total.actual);
      document.getElementById("kpiActualDelta").textContent = t().kpis.overCheck(total.over, total.check);
      document.getElementById("kpiGap").textContent = signedUsd(total.gap);
      document.getElementById("kpiGapDelta").textContent = total.gap < 0 ? t().kpis.shortfall : t().kpis.overCollected;
      document.getElementById("kpiBl").textContent = num(total.bl);
      document.getElementById("kpiBlDelta").textContent = t().kpis.rows(sourceRows.length);
      document.getElementById("kpiTeu").textContent = num(total.teu);
      document.getElementById("kpiTeuDelta").textContent = `${state.program === "ALL" ? "EFC + LSS" : state.program}`;
    }

    function renderBreadcrumb() {
      const bc = document.getElementById("breadcrumb");
      const parts = [`<button data-action="reset">${escapeHtml(t().buttons.all)}</button>`];
      if (state.selectedOrigin) {
        parts.push(`<span>/</span><button data-action="origin">${escapeHtml(state.selectedOrigin)}</button>`);
      }
      if (state.selectedDestination) {
        parts.push(`<span>/</span><button data-action="destination">${escapeHtml(state.selectedDestination)}</button>`);
      }
      bc.innerHTML = parts.join("");
      bc.querySelectorAll("button").forEach(button => {
        button.addEventListener("click", () => {
          const action = button.dataset.action;
          if (action === "reset") {
            state.selectedOrigin = "";
            state.selectedDestination = "";
            state.level = "origin";
          } else if (action === "origin") {
            state.selectedDestination = "";
            state.level = "lane";
          } else {
            state.level = "customer";
          }
          syncSegments();
          render();
        });
      });
    }

    function sortAggregates(items) {
      const key = state.tableSort.key;
      const direction = state.tableSort.direction === "asc" ? 1 : -1;
      return items.sort((a, b) => {
        let av = a[key];
        let bv = b[key];
        if (Array.isArray(av)) av = av.join(" > ");
        if (Array.isArray(bv)) bv = bv.join(" > ");
        if (typeof av === "string" || typeof bv === "string") return String(av).localeCompare(String(bv)) * direction;
        if (!Number.isFinite(av)) av = -Infinity;
        if (!Number.isFinite(bv)) bv = -Infinity;
        return (av - bv) * direction;
      });
    }

    function headerCell(label, key, numeric = false) {
      const cls = numeric ? " class=\"num\"" : "";
      return `<th${cls} data-sort="${key}">${label}</th>`;
    }

    function renderMainTable(sourceRows) {
      const groupers = currentGroupers();
      const items = sortAggregates(aggregate(sourceRows, groupers));
      const table = document.getElementById("mainTable");
      const title = document.getElementById("mainTitle");
      title.textContent = lang === "ko"
        ? `${t().labels[state.level]}${t().table.titleSuffix}`
        : `${t().labels[state.level]}${t().table.titleSuffix}`;

      if (!items.length) {
        table.innerHTML = `<tbody><tr><td class="empty">${escapeHtml(t().table.noData)}</td></tr></tbody>`;
        return;
      }

      const nameHeaders = state.level === "origin"
        ? [headerCell(t().labels.origin, "name")]
        : state.level === "lane"
          ? [headerCell(t().labels.origin, "origin"), headerCell(t().filters.destination, "destination")]
          : state.level === "salesperson"
            ? [headerCell(t().labels.origin, "origin"), headerCell(t().filters.destination, "destination"), headerCell(t().labels.salesperson, "salesperson")]
            : [headerCell(t().labels.origin, "origin"), headerCell(t().filters.destination, "destination"), headerCell(t().labels.customer, "customer"), headerCell(t().labels.salesperson, "salesText")];

      table.innerHTML = `
        <thead>
          <tr>
            ${nameHeaders.join("")}
            ${headerCell(t().table.bl, "bl", true)}
            ${headerCell(t().table.teu, "teu", true)}
            ${headerCell(t().labels.expected, "expected", true)}
            ${headerCell(t().labels.actual, "actual", true)}
            ${headerCell(t().labels.gap, "gap", true)}
            ${headerCell(t().labels.rate, "rate", true)}
            ${headerCell(t().table.issue, "missing", true)}
            <th>${escapeHtml(t().table.category)}</th>
          </tr>
        </thead>
        <tbody>
          ${items.map(item => {
            const cells = state.level === "origin"
              ? `<td>${escapeHtml(item.labels[0])}</td>`
              : state.level === "lane"
                ? `<td>${escapeHtml(item.labels[0])}</td><td>${escapeHtml(item.labels[1])}</td>`
                : state.level === "salesperson"
                  ? `<td>${escapeHtml(item.labels[0])}</td><td>${escapeHtml(item.labels[1])}</td><td>${escapeHtml(item.labels[2])}</td>`
                  : `<td>${escapeHtml(item.labels[0])}</td><td>${escapeHtml(item.labels[1])}</td><td>${escapeHtml(item.labels[2])}</td><td>${escapeHtml(safe(item.salesText))}</td>`;
            const issueCount = item.missing + item.partial;
            return `
              <tr class="data-row" data-parts="${escapeHtml(JSON.stringify(item.parts))}">
                ${cells}
                <td class="num">${num(item.bl)}</td>
                <td class="num">${num(item.teu)}</td>
                <td class="num">${usd(item.expected)}</td>
                <td class="num">${usd(item.actual)}</td>
                <td class="num">${signedUsd(item.gap)}</td>
                <td class="num">${pct(item.rate)}</td>
                <td class="num">${num(issueCount)}</td>
                <td>${escapeHtml(item.programText)}</td>
              </tr>
            `;
          }).join("")}
        </tbody>
      `;

      table.querySelectorAll("th[data-sort]").forEach(th => {
        th.addEventListener("click", () => {
          const key = th.dataset.sort;
          const map = { name: "sortLabel", origin: "sortLabel", destination: "sortLabel", customer: "sortLabel", salesperson: "sortLabel" };
          state.tableSort.key = map[key] || key;
          state.tableSort.direction = state.tableSort.direction === "desc" ? "asc" : "desc";
          render();
        });
      });

      table.querySelectorAll("tr.data-row").forEach(tr => {
        tr.addEventListener("click", () => {
          const parts = JSON.parse(tr.dataset.parts);
          if (state.level === "origin") {
            state.selectedOrigin = parts[0];
            state.selectedDestination = "";
            state.level = "lane";
          } else if (state.level === "lane") {
            state.selectedOrigin = parts[0];
            state.selectedDestination = parts[1];
            state.level = "customer";
          }
          syncSegments();
          render();
        });
      });
    }

    function renderSalesStatus(sourceRows) {
      const panel = document.getElementById("salesStatus");
      const groupers = [{
        value: row => salespersonValue(row),
        label: row => salespersonValue(row),
      }];
      const items = aggregate(sourceRows, groupers)
        .sort((a, b) => {
          const aShortfall = Math.min(a.gap, 0);
          const bShortfall = Math.min(b.gap, 0);
          if (aShortfall !== bShortfall) return aShortfall - bShortfall;
          return b.issueShippers - a.issueShippers;
        })
        .slice(0, 12);

      if (!items.length) {
        panel.innerHTML = `<div class="empty">${escapeHtml(t().table.noData)}</div>`;
        return;
      }

      const mappedShippers = new Set(sourceRows
        .filter(row => row.salesperson && row.bookingShipper)
        .map(row => row.bookingShipper)).size;
      const ownerCount = new Set(sourceRows.map(row => row.salesperson).filter(Boolean)).size;
      const totalGap = sourceRows.reduce((sum, row) => sum + Number(row.gap || 0), 0);
      const maxShortfall = Math.max(...items.map(item => Math.abs(Math.min(item.gap, 0))), 1);

      panel.innerHTML = `
        <div class="sales-status-summary">
          <div><span>${escapeHtml(t().table.mappedShippers)}</span><strong>${num(mappedShippers)}</strong></div>
          <div><span>${escapeHtml(t().table.owners)}</span><strong>${num(ownerCount)}</strong></div>
          <div><span>${escapeHtml(t().labels.gap)}</span><strong>${signedUsd(totalGap)}</strong></div>
        </div>
        <table class="sales-status-table">
          <thead>
            <tr>
              <th>${escapeHtml(t().table.salespersonStatus)}</th>
              <th class="num">${escapeHtml(t().table.shipperCount)}</th>
              <th class="num">${escapeHtml(t().table.missingShipperCount)}</th>
              <th class="num">${escapeHtml(t().table.partialShipperCount)}</th>
              <th class="num">${escapeHtml(t().labels.expected)}</th>
              <th class="num">${escapeHtml(t().labels.actual)}</th>
              <th class="num">${escapeHtml(t().labels.gap)}</th>
              <th class="num">${escapeHtml(t().labels.rate)}</th>
            </tr>
          </thead>
          <tbody>
            ${items.map(item => {
              const displayGap = item.gap;
              const shortfallWidth = Math.min(100, Math.abs(Math.min(displayGap, 0)) / maxShortfall * 100);
              return `
                <tr>
                  <td class="sales-name" title="${escapeHtml(item.labels[0])}">${escapeHtml(item.labels[0])}</td>
                  <td class="num">${num(item.shippers)}</td>
                  <td class="num">${num(item.missingShippers)}</td>
                  <td class="num">${num(item.partialShippers)}</td>
                  <td class="num">${usd(item.expected)}</td>
                  <td class="num">${usd(item.actual)}</td>
                  <td class="num sales-gap-cell">
                    <div class="sales-gap-text">${signedUsd(displayGap)}</div>
                    <div class="sales-gap-bar"><span style="width:${shortfallWidth}%"></span></div>
                  </td>
                  <td class="num">${pct(item.rate)}</td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>
      `;
    }

    function renderExceptions(sourceRows) {
      const table = document.getElementById("exceptionTable");
      const items = sourceRows
        .filter(row => !["정상", "수량없음"].includes(row.status))
        .sort((a, b) => Math.abs(b.gap) - Math.abs(a.gap))
        .slice(0, 120);
      if (!items.length) {
        table.innerHTML = `<tbody><tr><td class="empty">${escapeHtml(t().table.noException)}</td></tr></tbody>`;
        return;
      }
      table.innerHTML = `
        <thead>
          <tr>
            <th>${escapeHtml(t().table.status)}</th><th>${escapeHtml(t().table.bl)}</th><th>${escapeHtml(t().table.bookingShipper)}</th><th>${escapeHtml(t().table.salesperson)}</th><th>${escapeHtml(t().table.charge)}</th><th>${escapeHtml(t().table.pol)}</th><th>${escapeHtml(t().table.pod)}</th><th>${escapeHtml(t().labels.customer)}</th>
            <th class="num">20</th><th class="num">40</th><th class="num">${escapeHtml(t().labels.expected)}</th><th class="num">${escapeHtml(t().labels.actual)}</th><th class="num">${escapeHtml(t().labels.gap)}</th><th>${escapeHtml(t().table.tariffCat)}</th>
          </tr>
        </thead>
        <tbody>
          ${items.map(row => `
            <tr>
              <td><span class="pill ${statusClass(row.status)}">${escapeHtml(statusText(row.status))}</span></td>
              <td>${escapeHtml(row.bl)}</td>
              <td>${escapeHtml(safe(row.bookingShipper))}</td>
              <td>${escapeHtml(salespersonValue(row))}</td>
              <td>${escapeHtml(row.program)}</td>
              <td>${escapeHtml(row.originPort)} (${escapeHtml(row.originCountry)})</td>
              <td>${escapeHtml(row.destinationPort)} (${escapeHtml(row.destinationCountry)})</td>
              <td>${escapeHtml(customerValue(row))}</td>
              <td class="num">${num(row.qty20)}</td>
              <td class="num">${num(row.qty40)}</td>
              <td class="num">${usd(row.expected)}</td>
              <td class="num">${usd(row.actual)}</td>
              <td class="num">${signedUsd(row.gap)}</td>
              <td>${escapeHtml(row.tariffCategory)}</td>
            </tr>
          `).join("")}
        </tbody>
      `;
    }

    function syncSegments() {
      document.querySelectorAll("[data-segment]").forEach(group => {
        const key = group.dataset.segment;
        group.querySelectorAll("button").forEach(button => {
          button.classList.toggle("active", button.dataset.value === state[key]);
        });
      });
    }

    function render() {
      updateOriginDestinationFilters();
      const sourceRows = filteredRows();
      renderKpis(sourceRows);
      renderBreadcrumb();
      renderMainTable(sourceRows);
      renderSalesStatus(sourceRows);
      renderExceptions(sourceRows);
      document.getElementById("sourceMeta").textContent = t().sourceMeta(meta);
    }

    document.querySelectorAll("[data-segment] button").forEach(button => {
      button.addEventListener("click", () => {
        const group = button.closest("[data-segment]");
        const key = group.dataset.segment;
        state[key] = button.dataset.value;
        if (key === "originBasis") {
          state.origin = "ALL";
          state.selectedOrigin = "";
          state.selectedDestination = "";
        }
        if (key === "program") {
          state.origin = "ALL";
          state.destination = "ALL";
          state.salesperson = "ALL";
          state.selectedOrigin = "";
          state.selectedDestination = "";
        }
        if (key === "level" && state.level === "origin") {
          state.selectedOrigin = "";
          state.selectedDestination = "";
        }
        state.tableSort.key = state.sortMetric;
        if (key === "sortMetric") {
          state.tableSort.direction = state.sortMetric === "gap" ? "asc" : "desc";
        }
        syncSegments();
        render();
      });
    });

    document.getElementById("monthFilter").addEventListener("change", event => { state.month = event.target.value; render(); });
    document.getElementById("weekFilter").addEventListener("change", event => { state.week = event.target.value; render(); });
    document.getElementById("pcFilter").addEventListener("change", event => { state.pc = event.target.value; render(); });
    document.getElementById("statusFilter").addEventListener("change", event => { state.status = event.target.value; render(); });
    document.getElementById("salespersonFilter").addEventListener("change", event => { state.salesperson = event.target.value; render(); });
    document.getElementById("originFilter").addEventListener("change", event => {
      state.origin = event.target.value;
      state.selectedOrigin = "";
      state.selectedDestination = "";
      render();
    });
    document.getElementById("destinationFilter").addEventListener("change", event => {
      state.destination = event.target.value;
      state.selectedDestination = "";
      render();
    });
    document.getElementById("searchFilter").addEventListener("input", event => {
      state.search = event.target.value;
      render();
    });

    applyLang();
    checkSession();
  </script>
</body>
</html>
"""


def main() -> None:
    if not SOURCE_CSV.exists():
        raise SystemExit(f"Missing source file: {SOURCE_CSV}")

    rows, meta = read_rows()
    data = json.dumps({"rows": rows, "meta": meta}, ensure_ascii=False, separators=(",", ":"))
    data = data.replace("</", "<\\/")
    DATA_FILE.write_text(data, encoding="utf-8")
    html = HTML
    for output_file in OUTPUT_FILES:
        output_file.write_text(html, encoding="utf-8")
    print(f"Wrote {', '.join(str(path) for path in (*OUTPUT_FILES, DATA_FILE))} with {len(rows):,} target rows")
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
