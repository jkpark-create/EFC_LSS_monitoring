from __future__ import annotations

import argparse
import csv
import datetime as dt
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


DEFAULT_DOCUMENT_NAME = "[영업팀] LSS & EFC 징수금액조회"
DEFAULT_OUTPUT_CSV = Path("DynamicList.CSV")
DEFAULT_DOWNLOAD_DIR = Path("downloads")
DEFAULT_BROWSER_PROFILE = Path(".icc-browser")
REPORT_WINDOW_WEEKS = 4
CSV_OUTPUT_ENCODING = "cp949"
CSV_REQUIRED_HEADERS = ("실적년", "실적월", "실적년주차")


@dataclass(frozen=True)
class ReportWindow:
    start_year: int
    start_week: int
    end_year: int
    end_week: int

    def as_log_text(self) -> str:
        return (
            f"시작년주 {self.start_year}{self.start_week:02d}, "
            f"종료년주 {self.end_year}{self.end_week:02d}"
        )


def env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "t", "yes", "y", "on"}


def env_int(name: str) -> int | None:
    value = os.getenv(name)
    if value is None or not value.strip():
        return None
    return int(value)


def parse_date(value: str | None) -> dt.date:
    if not value:
        return dt.date.today()
    return dt.date.fromisoformat(value)


def shift_iso_week(year: int, week: int, delta_weeks: int) -> tuple[int, int]:
    monday = dt.date.fromisocalendar(year, week, 1)
    shifted = monday + dt.timedelta(weeks=delta_weeks)
    iso = shifted.isocalendar()
    return iso.year, iso.week


def current_icc_week(today: dt.date) -> tuple[int, int]:
    # ICC's business week is one week behind Python's ISO calendar in this workflow.
    target_day = today - dt.timedelta(days=7)
    iso = target_day.isocalendar()
    return iso.year, iso.week


def report_window(
    today: dt.date,
    weeks: int,
    target_year: int | None = None,
    target_week: int | None = None,
    start_year: int | None = None,
    start_week: int | None = None,
) -> ReportWindow:
    if weeks < 1:
        raise ValueError("--weeks must be at least 1")

    if target_year is None and target_week is None:
        end_year, end_week = current_icc_week(today)
    elif target_year is not None and target_week is not None:
        dt.date.fromisocalendar(target_year, target_week, 1)
        end_year, end_week = target_year, target_week
    else:
        raise ValueError("--target-year and --target-week must be used together")

    if start_year is None and start_week is None:
        first_year, first_week = shift_iso_week(end_year, end_week, -(weeks - 1))
    elif start_year is not None and start_week is not None:
        dt.date.fromisocalendar(start_year, start_week, 1)
        first_year, first_week = start_year, start_week
    else:
        raise ValueError("--start-year and --start-week must be used together")

    return ReportWindow(first_year, first_week, end_year, end_week)


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def first_existing_file(path: Path) -> Path:
    if not path.exists():
        raise FileNotFoundError(path)
    if path.is_dir():
        files = [item for item in path.iterdir() if item.is_file()]
        if not files:
            raise FileNotFoundError(f"No files in {path}")
        return max(files, key=lambda item: item.stat().st_mtime)
    return path


def text_score(text: str) -> int:
    lower_text = text.lower()
    score = 0
    score += sum(100 for header in CSV_REQUIRED_HEADERS if header in text)
    score += sum(10 for token in ("<html", "<table") if token in lower_text)
    score -= text.count("\ufffd") * 100
    return score


def detect_text_encoding(path: Path, encodings: Iterable[str] = ("cp949", "utf-8-sig", "utf-8")) -> str:
    last_error: UnicodeDecodeError | None = None
    best_encoding: str | None = None
    best_score: int | None = None

    for encoding in encodings:
        try:
            text = path.read_text(encoding=encoding)
        except UnicodeDecodeError as exc:
            last_error = exc
            continue

        score = text_score(text)
        if best_score is None or score > best_score:
            best_encoding = encoding
            best_score = score

    if best_encoding:
        return best_encoding
    if last_error:
        raise last_error
    return "utf-8"


def write_rows_to_csv(rows: list[list[str]], output_csv: Path) -> None:
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with output_csv.open("w", encoding=CSV_OUTPUT_ENCODING, errors="replace", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(rows)


def normalize_csv_file(source: Path, output_csv: Path) -> None:
    encoding = detect_text_encoding(source)
    with source.open("r", encoding=encoding, newline="") as src:
        rows = list(csv.reader(src))
    write_rows_to_csv(rows, output_csv)


def xml_name(name: str) -> str:
    return f"{{http://schemas.openxmlformats.org/spreadsheetml/2006/main}}{name}"


def rel_name(name: str) -> str:
    return f"{{http://schemas.openxmlformats.org/package/2006/relationships}}{name}"


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []

    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for item in root.findall(xml_name("si")):
        texts = [node.text or "" for node in item.iter(xml_name("t"))]
        values.append("".join(texts))
    return values


def first_sheet_path(zf: zipfile.ZipFile) -> str:
    names = set(zf.namelist())
    if "xl/workbook.xml" not in names or "xl/_rels/workbook.xml.rels" not in names:
        return "xl/worksheets/sheet1.xml"

    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall(rel_name("Relationship"))
        if "Id" in rel.attrib and "Target" in rel.attrib
    }
    sheets = workbook.find(xml_name("sheets"))
    if sheets is None:
        return "xl/worksheets/sheet1.xml"

    for sheet in sheets.findall(xml_name("sheet")):
        rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rel_targets.get(rel_id or "")
        if target:
            return "xl/" + target.lstrip("/")

    return "xl/worksheets/sheet1.xml"


def column_index(cell_ref: str | None, fallback: int) -> int:
    if not cell_ref:
        return fallback
    match = re.match(r"([A-Z]+)", cell_ref.upper())
    if not match:
        return fallback

    index = 0
    for char in match.group(1):
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def xlsx_cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    value = cell.find(xml_name("v"))

    if cell_type == "s" and value is not None and value.text is not None:
        index = int(value.text)
        return shared_strings[index] if 0 <= index < len(shared_strings) else ""
    if cell_type == "inlineStr":
        texts = [node.text or "" for node in cell.iter(xml_name("t"))]
        return "".join(texts)
    if value is None or value.text is None:
        return ""
    return value.text


def xlsx_to_rows(source: Path) -> list[list[str]]:
    with zipfile.ZipFile(source) as zf:
        shared_strings = read_shared_strings(zf)
        sheet_path = first_sheet_path(zf)
        root = ET.fromstring(zf.read(sheet_path))

    rows: list[list[str]] = []
    for row in root.iter(xml_name("row")):
        values: list[str] = []
        for cell in row.findall(xml_name("c")):
            index = column_index(cell.attrib.get("r"), len(values))
            while len(values) <= index:
                values.append("")
            values[index] = xlsx_cell_text(cell, shared_strings)
        while values and values[-1] == "":
            values.pop()
        rows.append(values)
    return rows


class TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.tables: list[list[list[str]]] = []
        self._table_depth = 0
        self._current_table: list[list[str]] | None = None
        self._current_row: list[str] | None = None
        self._current_cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "table":
            self._table_depth += 1
            if self._table_depth == 1:
                self._current_table = []
        elif tag == "tr" and self._table_depth == 1:
            self._current_row = []
        elif tag in {"td", "th"} and self._table_depth == 1 and self._current_row is not None:
            self._current_cell = []
        elif tag == "br" and self._current_cell is not None:
            self._current_cell.append(" ")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in {"td", "th"} and self._current_cell is not None and self._current_row is not None:
            self._current_row.append(normalize_text("".join(self._current_cell)))
            self._current_cell = None
        elif tag == "tr" and self._current_row is not None and self._current_table is not None:
            if any(cell for cell in self._current_row):
                self._current_table.append(self._current_row)
            self._current_row = None
        elif tag == "table" and self._table_depth:
            if self._table_depth == 1 and self._current_table is not None:
                self.tables.append(self._current_table)
                self._current_table = None
            self._table_depth -= 1

    def handle_data(self, data: str) -> None:
        if self._current_cell is not None:
            self._current_cell.append(data)


def html_table_to_rows(source: Path) -> list[list[str]]:
    encoding = detect_text_encoding(source)
    parser = TableParser()
    parser.feed(source.read_text(encoding=encoding, errors="replace"))
    if not parser.tables:
        raise ValueError(f"No HTML table found in {source}")

    for table in parser.tables:
        header = {normalize_text(cell) for row in table[:3] for cell in row}
        if {"실적년", "실적월", "실적년주차"} <= header:
            return table
    return max(parser.tables, key=len)


def excel_com_to_csv(source: Path, output_csv: Path) -> None:
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError(
            "Binary .xls files require Excel COM support. "
            "Install pywin32 or change ICC to download CSV/XLSX."
        ) from exc

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(source.resolve()))
        workbook.SaveAs(str(output_csv.resolve()), FileFormat=6)
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()


def convert_download_to_csv(source: Path, output_csv: Path) -> None:
    source = first_existing_file(source)
    suffix = source.suffix.lower()

    if suffix in {".csv", ".txt"}:
        normalize_csv_file(source, output_csv)
    elif suffix == ".xlsx":
        write_rows_to_csv(xlsx_to_rows(source), output_csv)
    elif suffix == ".xls":
        sample = source.read_bytes()[:512].lstrip()
        if sample.startswith((b"<", b"\xef\xbb\xbf<")):
            write_rows_to_csv(html_table_to_rows(source), output_csv)
        else:
            excel_com_to_csv(source, output_csv)
    else:
        try:
            normalize_csv_file(source, output_csv)
        except UnicodeDecodeError as exc:
            raise ValueError(f"Unsupported download type: {source}") from exc


def field_selector(name: str) -> str | None:
    return os.getenv(f"ICC_SELECTOR_{name.upper()}")


def set_value_by_selector(page, selector: str, value: str, timeout: int) -> bool:
    locator = page.locator(selector).first
    locator.wait_for(state="visible", timeout=timeout)
    tag = locator.evaluate("element => element.tagName.toLowerCase()")
    if tag == "select":
        try:
            locator.select_option(value=value, timeout=timeout)
        except Exception:
            locator.select_option(label=value, timeout=timeout)
    else:
        locator.fill(value, timeout=timeout)
    return True


def set_values_near_label(page, label: str, values: list[str], timeout: int) -> bool:
    result = page.evaluate(
        """
        ([label, values]) => {
          const normalize = (value) => (value || "").replace(/\\s+/g, " ").trim();
          const visible = (element) => {
            const style = window.getComputedStyle(element);
            const box = element.getBoundingClientRect();
            return style.visibility !== "hidden" && style.display !== "none" && box.width >= 0 && box.height >= 0;
          };
          const editable = (element) => {
            return ["INPUT", "TEXTAREA", "SELECT"].includes(element.tagName) || element.isContentEditable;
          };
          const setValue = (element, value) => {
            if (element.tagName === "SELECT") {
              const options = Array.from(element.options || []);
              const option = options.find((item) => item.value === value || normalize(item.textContent) === value);
              if (option) element.value = option.value;
              else element.value = value;
            } else if (element.isContentEditable) {
              element.textContent = value;
            } else {
              element.value = value;
            }
            element.dispatchEvent(new Event("input", { bubbles: true }));
            element.dispatchEvent(new Event("change", { bubbles: true }));
            element.dispatchEvent(new Event("blur", { bubbles: true }));
          };
          const exactText = (element) => normalize(element.innerText || element.textContent) === label;
          const labels = Array.from(document.querySelectorAll("body *")).filter((element) => visible(element) && exactText(element));
          for (const labelElement of labels) {
            const scopes = [];
            let tr = labelElement.closest("tr");
            if (tr) scopes.push(tr);
            let parent = labelElement.parentElement;
            for (let index = 0; parent && index < 5; index += 1, parent = parent.parentElement) {
              scopes.push(parent);
            }
            for (const scope of scopes) {
              const fields = Array.from(scope.querySelectorAll("input, textarea, select, [contenteditable='true']"))
                .filter((element) => visible(element) && editable(element));
              if (fields.length >= values.length) {
                values.forEach((value, index) => setValue(fields[index], value));
                return true;
              }
            }
          }
          return false;
        }
        """,
        [label, values],
    )
    if result:
        return True

    try:
        page.get_by_text(label, exact=True).click(timeout=timeout)
        for value in values:
            page.keyboard.press("Tab")
            page.keyboard.press("Control+A")
            page.keyboard.type(value)
        return True
    except Exception:
        return False


def set_year_week_near_label(page, label: str, year: str, week: str) -> bool:
    return bool(
        page.evaluate(
            """
            ([label, year, week]) => {
              const compact = `${year}${String(week).padStart(2, "0")}`;
              const normalize = (value) => (value || "").replace(/\\s+/g, " ").trim();
              const visible = (element) => {
                const style = window.getComputedStyle(element);
                const box = element.getBoundingClientRect();
                return style.visibility !== "hidden" && style.display !== "none" && box.width >= 0 && box.height >= 0;
              };
              const editable = (element) => {
                return ["INPUT", "TEXTAREA", "SELECT"].includes(element.tagName) || element.isContentEditable;
              };
              const setValue = (element, value) => {
                if (element.tagName === "SELECT") {
                  const options = Array.from(element.options || []);
                  const option = options.find((item) => item.value === value || normalize(item.textContent) === value);
                  if (option) element.value = option.value;
                  else element.value = value;
                } else if (element.isContentEditable) {
                  element.textContent = value;
                } else {
                  element.value = value;
                }
                element.dispatchEvent(new Event("input", { bubbles: true }));
                element.dispatchEvent(new Event("change", { bubbles: true }));
                element.dispatchEvent(new Event("blur", { bubbles: true }));
              };
              const exactText = (element) => normalize(element.innerText || element.textContent) === label;
              const labels = Array.from(document.querySelectorAll("body *")).filter((element) => visible(element) && exactText(element));
              for (const labelElement of labels) {
                const scopes = [];
                let tr = labelElement.closest("tr");
                if (tr) scopes.push(tr);
                let parent = labelElement.parentElement;
                for (let index = 0; parent && index < 5; index += 1, parent = parent.parentElement) {
                  scopes.push(parent);
                }
                for (const scope of scopes) {
                  const fields = Array.from(scope.querySelectorAll("input, textarea, select, [contenteditable='true']"))
                    .filter((element) => visible(element) && editable(element));
                  if (fields.length === 1) {
                    setValue(fields[0], compact);
                    return true;
                  }
                  if (fields.length >= 2) {
                    setValue(fields[0], year);
                    setValue(fields[1], week);
                    return true;
                  }
                }
              }
              return false;
            }
            """,
            [label, year, week],
        )
    )


def set_named_field(page, name: str, label: str, values: list[str], timeout: int) -> None:
    compact_value = "".join(values)
    single_selector = field_selector(name)
    selectors = [single_selector]
    if len(values) == 2:
        split_selectors = [field_selector(f"{name}_YEAR"), field_selector(f"{name}_WEEK")]
        if all(split_selectors):
            selectors = split_selectors

    if selectors and all(selectors):
        selector_values = values if len(selectors) > 1 else [compact_value]
        for selector, value in zip(selectors, selector_values, strict=True):
            set_value_by_selector(page, selector or "", value, timeout)
        return

    if len(values) == 2 and set_year_week_near_label(page, label, values[0], values[1]):
        return

    if set_values_near_label(page, label, values, timeout):
        return

    raise RuntimeError(
        f"Could not set ICC field '{label}'. "
        f"Set ICC_SELECTOR_{name.upper()} or related selector environment variables."
    )


def select_document(page, document_name: str, timeout: int) -> None:
    selector = field_selector("DOCUMENT")
    if selector:
        set_value_by_selector(page, selector, document_name, timeout)
        return

    selected = page.evaluate(
        """
        (documentName) => {
          const normalize = (value) => (value || "").replace(/\\s+/g, " ").trim();
          for (const select of Array.from(document.querySelectorAll("select"))) {
            const option = Array.from(select.options || [])
              .find((item) => normalize(item.textContent) === documentName || item.value === documentName);
            if (option) {
              select.value = option.value;
              select.dispatchEvent(new Event("change", { bubbles: true }));
              return true;
            }
          }
          return false;
        }
        """,
        document_name,
    )
    if selected:
        return

    # Some ICC screens use a custom combo widget. This lets Playwright select it when text is visible.
    try:
        page.get_by_text(document_name, exact=True).click(timeout=timeout)
    except Exception as exc:
        raise RuntimeError(
            "Could not select Document Name. Set ICC_SELECTOR_DOCUMENT if the ICC combo is custom."
        ) from exc


def click_text_button(page, text: str, timeout: int) -> None:
    selector = field_selector(text.upper().replace(" ", "_"))
    if selector:
        page.locator(selector).first.click(timeout=timeout)
        return

    candidates = [
        lambda: page.get_by_role("button", name=re.compile(re.escape(text), re.I)).click(timeout=timeout),
        lambda: page.get_by_text(text, exact=True).click(timeout=timeout),
        lambda: page.locator(f"text={text}").first.click(timeout=timeout),
    ]
    last_error: Exception | None = None
    for candidate in candidates:
        try:
            candidate()
            return
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"Could not click ICC button/text '{text}'") from last_error


def wait_after_action(page, timeout: int) -> None:
    try:
        page.wait_for_load_state("networkidle", timeout=timeout)
    except Exception:
        time.sleep(3)


def download_from_icc(args: argparse.Namespace, window: ReportWindow) -> Path:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError(
            "Playwright is required for ICC browser automation. "
            "Run: py -m pip install -r requirements.txt && py -m playwright install chromium"
        ) from exc

    timeout = args.timeout * 1000
    download_dir = Path(args.download_dir)
    download_dir.mkdir(parents=True, exist_ok=True)
    profile_dir = Path(args.browser_profile)
    profile_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            str(profile_dir),
            headless=args.headless,
            accept_downloads=True,
            downloads_path=str(download_dir),
            slow_mo=args.slow_mo,
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(timeout)

        if args.url:
            page.goto(args.url, wait_until="domcontentloaded", timeout=timeout)
            wait_after_action(page, timeout)
        elif not page.url or page.url == "about:blank":
            raise RuntimeError("ICC URL is required on first run. Pass --url or set ICC_URL.")

        select_document(page, args.document_name, timeout)
        set_named_field(page, "START", "시작년주", [str(window.start_year), str(window.start_week)], timeout)
        set_named_field(page, "END", "종료년주", [str(window.end_year), str(window.end_week)], timeout)
        set_named_field(page, "ORG", "조직", [args.org], timeout)
        set_named_field(page, "DIVISION", "구분", [args.division], timeout)

        click_text_button(page, args.search_text, timeout)
        wait_after_action(page, timeout)

        with page.expect_download(timeout=timeout) as download_info:
            click_text_button(page, args.download_text, timeout)
        download = download_info.value
        timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        suggested = download.suggested_filename or "icc_download.xlsx"
        target = download_dir / f"{timestamp}_{suggested}"
        download.save_as(str(target))
        context.close()
        return target


def run_dashboard_build() -> None:
    subprocess.run([sys.executable, "build_dashboard.py"], check=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Download ICC data and rebuild the EFC/LSS dashboard.")
    parser.add_argument("--url", default=os.getenv("ICC_URL"), help="ICC On-Demand Data page URL.")
    parser.add_argument("--document-name", default=os.getenv("ICC_DOCUMENT_NAME", DEFAULT_DOCUMENT_NAME))
    parser.add_argument("--org", default=os.getenv("ICC_ORG", "O"), help="ICC 조직 value.")
    parser.add_argument("--division", default=os.getenv("ICC_DIVISION", "D"), help="ICC 구분 value.")
    parser.add_argument("--weeks", type=int, default=int(os.getenv("ICC_WINDOW_WEEKS", str(REPORT_WINDOW_WEEKS))))
    parser.add_argument("--date", default=os.getenv("ICC_RUN_DATE"), help="Override run date as YYYY-MM-DD.")
    parser.add_argument("--target-year", type=int, default=env_int("ICC_TARGET_YEAR"))
    parser.add_argument("--target-week", type=int, default=env_int("ICC_TARGET_WEEK"))
    parser.add_argument("--start-year", type=int, default=env_int("ICC_START_YEAR"))
    parser.add_argument("--start-week", type=int, default=env_int("ICC_START_WEEK"))
    parser.add_argument("--download-dir", default=os.getenv("ICC_DOWNLOAD_DIR", str(DEFAULT_DOWNLOAD_DIR)))
    parser.add_argument("--browser-profile", default=os.getenv("ICC_BROWSER_PROFILE", str(DEFAULT_BROWSER_PROFILE)))
    parser.add_argument("--output-csv", default=os.getenv("ICC_OUTPUT_CSV", str(DEFAULT_OUTPUT_CSV)))
    parser.add_argument("--download-file", help="Use an existing ICC Excel/CSV file instead of opening ICC.")
    parser.add_argument("--search-text", default=os.getenv("ICC_SEARCH_TEXT", "Search"))
    parser.add_argument("--download-text", default=os.getenv("ICC_DOWNLOAD_TEXT", "Excel Down"))
    parser.add_argument("--timeout", type=int, default=int(os.getenv("ICC_TIMEOUT_SECONDS", "90")))
    parser.add_argument("--slow-mo", type=int, default=int(os.getenv("ICC_SLOW_MO_MS", "0")))
    parser.add_argument("--headless", action="store_true", default=env_bool("ICC_HEADLESS", False))
    parser.add_argument("--no-build", action="store_true", help="Do not rebuild index.html after CSV update.")
    parser.add_argument("--dry-run", action="store_true", help="Print the computed ICC conditions only.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    today = parse_date(args.date)
    window = report_window(
        today=today,
        weeks=args.weeks,
        target_year=args.target_year,
        target_week=args.target_week,
        start_year=args.start_year,
        start_week=args.start_week,
    )

    print(f"ICC 조건: {window.as_log_text()}, 조직 {args.org}, 구분 {args.division}")
    print(f"Document Name: {args.document_name}")

    if args.dry_run:
        return

    output_csv = Path(args.output_csv)
    if args.download_file:
        downloaded_file = first_existing_file(Path(args.download_file))
        print(f"Using existing download: {downloaded_file}")
    else:
        downloaded_file = download_from_icc(args, window)
        print(f"Downloaded ICC file: {downloaded_file}")

    backup_path = output_csv.with_suffix(output_csv.suffix + ".bak")
    if output_csv.exists():
        shutil.copy2(output_csv, backup_path)
        print(f"Backed up previous CSV: {backup_path}")

    convert_download_to_csv(downloaded_file, output_csv)
    print(f"Updated source CSV: {output_csv}")

    if not args.no_build:
        run_dashboard_build()


if __name__ == "__main__":
    main()
