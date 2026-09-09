"""
NPP Daily Coal Stock Pipeline

One self-contained class that:
1. Downloads the latest NPP Daily Coal Report.
2. Automatically tries XLSX first, then XLS.
3. Handles the three NPP layouts observed in the reports:
      - XLS_OLD : L/O/R/Z
      - XLS_NEW : M/R/Y with Daily derived as R/M
      - XLSX_STANDARD : H/I/L with Days derived as I/H
4. Finds the Grand Total row dynamically.
5. Validates the extracted numbers.
6. Upserts the result into a CSV (safe to run every day).
7. Can backfill any date range.
8. Creates an audit CSV with layout, source columns and validation results.

Dependencies:
    pip install requests pandas xlrd openpyxl

Daily use:
    pipeline = NPPCoalStockPipeline()
    pipeline.run_daily()

Backfill:
    pipeline.run_backfill("2026-08-01", "2026-09-07")

The default output folder is ./npp_coal_stock_data.
"""

from __future__ import annotations

import math
import re
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import requests
import urllib3


urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


@dataclass
class ExtractionResult:
    report_date: str
    file_name: str
    sheet_name: str
    layout: str
    grand_total_row: int

    normative_days: float
    daily_requirement: float
    normative_stock: float
    actual_stock_total: float

    normative_days_column: Optional[str]
    daily_column: Optional[str]
    normative_stock_column: str
    indigenous_column: Optional[str]
    import_column: Optional[str]
    actual_total_column: str

    days_method: str
    daily_method: str

    normative_validation_error_pct: float
    actual_validation_error_pct: Optional[float]
    status: str = "SUCCESS"

    def output_dict(self) -> dict[str, Any]:
        return {
            "Date": self.report_date,
            "Normative Stock Reqd. (Days)": self.normative_days,
            "Daily Requirement @85% PLF (In '000 Tonnes)": self.daily_requirement,
            "Actual Stock - Total (In '000 Tonnes)": self.actual_stock_total,
        }

    def audit_dict(self) -> dict[str, Any]:
        d = self.output_dict()
        d.update(
            {
                "_File": self.file_name,
                "_Sheet": self.sheet_name,
                "_Layout": self.layout,
                "_Grand_Total_Row": self.grand_total_row,
                "_Normative_Days_Column": self.normative_days_column,
                "_Daily_Column": self.daily_column or "DERIVED",
                "_Normative_Stock_Column": self.normative_stock_column,
                "_Indigenous_Column": self.indigenous_column,
                "_Import_Column": self.import_column,
                "_Actual_Total_Column": self.actual_total_column,
                "_Days_Method": self.days_method,
                "_Daily_Method": self.daily_method,
                "_Normative_Validation_Error_Pct": self.normative_validation_error_pct,
                "_Actual_Validation_Error_Pct": self.actual_validation_error_pct,
                "_Status": self.status,
            }
        )
        return d


class NPPCoalStockPipeline:
    """
    Fully self-contained NPP coal stock downloader + extractor + CSV writer.

    The class is designed to be run repeatedly. It does not depend on a
    manually supplied date for daily operation: run_daily() automatically
    targets yesterday, because NPP's daily report is normally published for
    the previous reporting day.
    """

    BASE_URL = "https://npp.gov.in/public-reports/cea/daily/fuel"

    OUTPUT_COLUMNS = [
        "Date",
        "Normative Stock Reqd. (Days)",
        "Daily Requirement @85% PLF (In '000 Tonnes)",
        "Actual Stock - Total (In '000 Tonnes)",
    ]

    # Known physical layouts in NPP reports.
    # Columns are zero-based indexes.
    XLS_OLD = {
        "days": 11,          # L
        "daily": 14,         # O
        "normative_stock": 17,# R
        "indigenous": 21,    # V
        "import": 23,        # X
        "actual_total": 25,  # Z
    }

    XLS_NEW = {
        "days": 12,          # M
        "normative_stock": 17,# R
        "indigenous": 18,    # S
        "import": 21,        # V
        "actual_total": 24,  # Y
    }

    XLSX_STANDARD = {
        "daily": 7,           # H
        "normative_stock": 8, # I
        "indigenous": 9,      # J
        "import": 10,        # K
        "actual_total": 11,   # L
    }

    def __init__(
        self,
        output_dir: str | Path = "npp_coal_stock_data",
        timeout: tuple[int, int] = (20, 90),
        retries: int = 3,
        request_pause: float = 0.5,
        overwrite_downloads: bool = False,
    ):
        self.output_dir = Path(output_dir)
        self.report_dir = self.output_dir / "reports"
        self.output_csv = self.output_dir / "npp_coal_stocks.csv"
        self.audit_csv = self.output_dir / "npp_coal_stocks_audit.csv"
        self.failed_csv = self.output_dir / "npp_coal_stocks_failures.csv"

        self.timeout = timeout
        self.retries = max(1, int(retries))
        self.request_pause = max(0.0, float(request_pause))
        self.overwrite_downloads = overwrite_downloads

        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.report_dir.mkdir(parents=True, exist_ok=True)

        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                    "AppleWebKit/537.36 Chrome/140 Safari/537.36"
                ),
                "Accept": "*/*",
                "Connection": "keep-alive",
            }
        )

    # ============================================================
    # GENERAL HELPERS
    # ============================================================

    @staticmethod
    def _norm(value: Any) -> str:
        if value is None:
            return ""

        try:
            if pd.isna(value):
                return ""
        except Exception:
            pass

        text = str(value).lower()
        text = (
            text.replace("\n", " ")
            .replace("\r", " ")
            .replace("\t", " ")
            .replace("’", "'")
            .replace("‘", "'")
            .replace("–", "-")
            .replace("—", "-")
        )
        return re.sub(r"\s+", " ", text).strip()

    @staticmethod
    def _is_num(value: Any) -> bool:
        if value is None:
            return False

        try:
            if pd.isna(value):
                return False
        except Exception:
            pass

        try:
            return math.isfinite(float(value))
        except Exception:
            return False

    @classmethod
    def _num(cls, value: Any) -> Optional[float]:
        return float(value) if cls._is_num(value) else None

    @staticmethod
    def _clean_num(value: float) -> float | int:
        value = float(value)
        if abs(value - round(value)) < 1e-9:
            return int(round(value))
        return round(value, 6)

    @staticmethod
    def _excel_col(index: int) -> str:
        result = ""
        n = index + 1
        while n:
            n, rem = divmod(n - 1, 26)
            result = chr(65 + rem) + result
        return result

    @staticmethod
    def _date_from_filename(filename: str) -> str:
        match = re.search(r"(20\d{2}-\d{2}-\d{2})", filename)
        if not match:
            raise ValueError(f"Date not found in filename: {filename}")
        return match.group(1)

    @staticmethod
    def _error_pct(actual: float, expected: float) -> float:
        return (
            abs(actual - expected) / max(abs(expected), 1.0) * 100.0
        )

    # ============================================================
    # URL / DOWNLOAD
    # ============================================================

    def create_urls(self, report_date: date) -> tuple[str, str]:
        day = f"{report_date.day:02d}"
        month = f"{report_date.month:02d}"
        year = report_date.year

        folder = f"{day}-{month}-{year}"

        xlsx_url = (
            f"{self.BASE_URL}/{folder}/"
            f"dailyCoal1-{year}-{month}-{day}.xlsx"
        )

        xls_url = (
            f"{self.BASE_URL}/{folder}/"
            f"dailyCoal1-{year}-{month}-{day}.xls"
        )

        return xlsx_url, xls_url

    def _download_url(self, url: str, destination: Path) -> bool:
        last_error = None

        for attempt in range(1, self.retries + 1):
            try:
                response = self.session.get(
                    url,
                    timeout=self.timeout,
                    verify=False,
                )

                if response.status_code == 200 and response.content:
                    # NPP occasionally returns an HTML error page with 200.
                    # Reject obvious HTML responses.
                    prefix = response.content[:100].lstrip().lower()
                    if prefix.startswith(b"<html") or b"<html" in prefix:
                        raise RuntimeError("NPP returned HTML instead of Excel")

                    destination.write_bytes(response.content)
                    return True

                last_error = (
                    f"HTTP {response.status_code} "
                    f"for {url}"
                )

            except Exception as exc:
                last_error = str(exc)

            if attempt < self.retries:
                time.sleep(self.request_pause * attempt)

        print(f"  Download failed: {last_error}")
        return False

    def download_report(self, report_date: date) -> Optional[Path]:
        """
        Download one report.

        XLSX is attempted first. If unavailable, XLS is attempted.
        Existing local files are reused unless overwrite_downloads=True.
        """
        date_string = report_date.strftime("%Y-%m-%d")
        xlsx_url, xls_url = self.create_urls(report_date)

        print(f"\n[{date_string}] Downloading NPP report...")

        # Prefer XLSX.
        xlsx_path = self.report_dir / f"dailyCoal1-{date_string}.xlsx"

        if xlsx_path.exists() and not self.overwrite_downloads:
            print(f"  Using existing: {xlsx_path.name}")
            return xlsx_path

        print("  Trying XLSX...")
        if self._download_url(xlsx_url, xlsx_path):
            print(f"  ✓ Downloaded: {xlsx_path.name}")
            return xlsx_path

        # Fallback XLS.
        xls_path = self.report_dir / f"dailyCoal1-{date_string}.xls"

        if xls_path.exists() and not self.overwrite_downloads:
            print(f"  Using existing: {xls_path.name}")
            return xls_path

        print("  XLSX unavailable. Trying XLS...")
        if self._download_url(xls_url, xls_path):
            print(f"  ✓ Downloaded: {xls_path.name}")
            return xls_path

        print(f"  ✗ No NPP report available for {date_string}")
        return None

    # ============================================================
    # WORKBOOK LOADING
    # ============================================================

    def _load_xls(self, path: Path) -> tuple[str, list[list[Any]]]:
        import xlrd

        wb = xlrd.open_workbook(path)

        best_sheet = None
        best_score = -1

        for sheet in wb.sheets():
            score = 0

            for r in range(min(sheet.nrows, 120)):
                for c in range(min(sheet.ncols, 50)):
                    text = self._norm(sheet.cell_value(r, c))

                    if "grand total" in text:
                        score += 100
                    if "daily requirement" in text:
                        score += 40
                    if "actual stock" in text:
                        score += 40
                    if "normative stock" in text:
                        score += 40

            # Prefer the expected sheet name when available.
            if self._norm(sheet.name) == "dailycoalreport":
                score += 1000

            if score > best_score:
                best_score = score
                best_sheet = sheet

        if best_sheet is None:
            raise RuntimeError(
                f"Could not find DailyCoalReport sheet in {path.name}"
            )

        values = [
            [
                best_sheet.cell_value(r, c)
                for c in range(best_sheet.ncols)
            ]
            for r in range(best_sheet.nrows)
        ]

        return best_sheet.name, values

    def _load_xlsx(self, path: Path) -> tuple[str, list[list[Any]]]:
        from openpyxl import load_workbook

        wb = load_workbook(
            path,
            read_only=True,
            data_only=True,
        )

        best_ws = None
        best_score = -1

        # Only inspect first 40 columns for detection. NPP XLSX files can
        # report enormous max_column values because of formatting.
        for ws in wb.worksheets:
            score = 0

            for row in ws.iter_rows(
                min_row=1,
                max_row=min(ws.max_row or 1, 120),
                min_col=1,
                max_col=min(ws.max_column or 1, 40),
                values_only=True,
            ):
                for value in row:
                    text = self._norm(value)

                    if "grand total" in text:
                        score += 100
                    if "daily requirement" in text:
                        score += 40
                    if "actual stock" in text:
                        score += 40
                    if "normative stock" in text:
                        score += 40

            if self._norm(ws.title) == "dailycoalreport":
                score += 1000

            if score > best_score:
                best_score = score
                best_ws = ws

        if best_ws is None:
            raise RuntimeError(
                f"Could not find DailyCoalReport sheet in {path.name}"
            )

        values = [
            list(row)
            for row in best_ws.iter_rows(values_only=True)
        ]

        return best_ws.title, values

    def load_workbook(self, path: Path) -> tuple[str, list[list[Any]]]:
        suffix = path.suffix.lower()

        if suffix == ".xls":
            return self._load_xls(path)

        if suffix == ".xlsx":
            return self._load_xlsx(path)

        raise ValueError(f"Unsupported workbook format: {path}")

    # ============================================================
    # ROW DETECTION
    # ============================================================

    def find_grand_total_row(self, values: list[list[Any]]) -> int:
        candidates: list[tuple[int, int]] = []

        for r, row in enumerate(values):
            text_parts = [
                self._norm(v)
                for v in row
                if self._norm(v)
            ]

            text = " ".join(text_parts)
            compact = (
                text
                .replace(" ", "")
                .replace(":", "")
                .replace("-", "")
            )

            score = 0

            if "grandtotal" in compact:
                score += 1000

            if "a+b+c+d" in compact:
                score += 1000

            if "कुलयोग" in compact:
                score += 500

            numeric_count = sum(self._is_num(v) for v in row)
            score += min(numeric_count, 30)

            if score:
                candidates.append((score, r))

        if not candidates:
            raise RuntimeError("Grand Total row not found")

        # Highest score wins. If tied, use the later row, because the
        # Grand Total is normally below the state/plant rows.
        candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)

        return candidates[0][1]

    # ============================================================
    # LAYOUT DETECTION
    # ============================================================

    @staticmethod
    def _get(values: list[list[Any]], row: int, col: int) -> Any:
        if row < 0 or row >= len(values):
            return None
        if col < 0 or col >= len(values[row]):
            return None
        return values[row][col]

    def detect_layout(
        self,
        path: Path,
        values: list[list[Any]],
        grand_total_row: int,
    ) -> str:
        suffix = path.suffix.lower()

        if suffix == ".xlsx":
            row = values[grand_total_row]

            daily = self._num(self._get(values, grand_total_row, 7))
            normative = self._num(self._get(values, grand_total_row, 8))
            actual = self._num(self._get(values, grand_total_row, 11))

            if (
                daily is not None
                and normative is not None
                and actual is not None
                and daily > 100
                and normative > 10000
                and actual > 0
            ):
                return "XLSX_STANDARD"

            raise RuntimeError(
                "Unknown XLSX NPP layout: expected H/I/L values "
                "were not found in Grand Total row."
            )

        if suffix == ".xls":
            # Test the old layout first using the strongest signature:
            # L = days, R = normative stock, Z = actual total.
            old_days = self._num(self._get(values, grand_total_row, 11))
            old_norm = self._num(self._get(values, grand_total_row, 17))
            old_actual = self._num(self._get(values, grand_total_row, 25))

            if (
                old_days is not None
                and old_norm is not None
                and old_actual is not None
                and 1 < old_days < 100
                and old_norm > 10000
                and old_actual > 0
            ):
                return "XLS_OLD"

            # New XLS:
            # M = days, R = normative stock, Y = actual total.
            new_days = self._num(self._get(values, grand_total_row, 12))
            new_norm = self._num(self._get(values, grand_total_row, 17))
            new_actual = self._num(self._get(values, grand_total_row, 24))

            if (
                new_days is not None
                and new_norm is not None
                and new_actual is not None
                and 1 < new_days < 100
                and new_norm > 10000
                and new_actual > 0
            ):
                return "XLS_NEW"

        raise RuntimeError(
            f"Unknown NPP workbook layout: {path.name}"
        )

    # ============================================================
    # EXTRACTION
    # ============================================================

    def _extract_xls_old(
        self,
        values: list[list[Any]],
        row: int,
    ) -> dict[str, Any]:
        c = self.XLS_OLD

        days = self._num(self._get(values, row, c["days"]))
        daily_reported = self._num(self._get(values, row, c["daily"]))
        normative_stock = self._num(
            self._get(values, row, c["normative_stock"])
        )
        indigenous = self._num(
            self._get(values, row, c["indigenous"])
        )
        imported = self._num(
            self._get(values, row, c["import"])
        )
        actual_total = self._num(
            self._get(values, row, c["actual_total"])
        )

        if days is None or not (0 < days < 100):
            raise RuntimeError(f"Invalid XLS_OLD Normative Days: {days}")

        if normative_stock is None or normative_stock <= 0:
            raise RuntimeError(
                f"Invalid XLS_OLD Normative Stock: {normative_stock}"
            )

        if actual_total is None or actual_total < 0:
            raise RuntimeError(
                f"Invalid XLS_OLD Actual Total: {actual_total}"
            )

        derived_daily = normative_stock / days

        # Old reports have a Daily Requirement cell at O. Prefer it when
        # it is numerically consistent; otherwise derive from the invariant.
        if daily_reported is not None and daily_reported > 0:
            daily_error = self._error_pct(
                daily_reported,
                derived_daily,
            )

            if daily_error <= 3.0:
                daily = daily_reported
                daily_method = "report_cell_validated"
            else:
                daily = derived_daily
                daily_method = "derived_due_to_mismatch"
        else:
            daily = derived_daily
            daily_method = "derived"

        actual_expected = None
        actual_error = None

        if indigenous is not None and imported is not None:
            actual_expected = indigenous + imported
            actual_error = self._error_pct(
                actual_total,
                actual_expected,
            )

        return {
            "days": days,
            "daily": daily,
            "normative_stock": normative_stock,
            "actual_total": actual_total,
            "indigenous": indigenous,
            "imported": imported,
            "actual_expected": actual_expected,
            "actual_error": actual_error,
            "days_method": "explicit",
            "daily_method": daily_method,
            "normative_days_col": self._excel_col(c["days"]),
            "daily_col": self._excel_col(c["daily"]),
            "normative_stock_col": self._excel_col(c["normative_stock"]),
            "indigenous_col": self._excel_col(c["indigenous"]),
            "import_col": self._excel_col(c["import"]),
            "actual_total_col": self._excel_col(c["actual_total"]),
        }

    def _extract_xls_new(
        self,
        values: list[list[Any]],
        row: int,
    ) -> dict[str, Any]:
        c = self.XLS_NEW

        days = self._num(self._get(values, row, c["days"]))
        normative_stock = self._num(
            self._get(values, row, c["normative_stock"])
        )
        indigenous = self._num(
            self._get(values, row, c["indigenous"])
        )
        imported = self._num(
            self._get(values, row, c["import"])
        )
        actual_total = self._num(
            self._get(values, row, c["actual_total"])
        )

        if days is None or not (0 < days < 100):
            raise RuntimeError(f"Invalid XLS_NEW Normative Days: {days}")

        if normative_stock is None or normative_stock <= 0:
            raise RuntimeError(
                f"Invalid XLS_NEW Normative Stock: {normative_stock}"
            )

        if actual_total is None or actual_total < 0:
            raise RuntimeError(
                f"Invalid XLS_NEW Actual Total: {actual_total}"
            )

        # In XLS_NEW the Grand Total Daily Requirement cells are blank.
        # The report's defining relationship is:
        #
        # Normative Stock = Daily Requirement × Normative Days
        #
        daily = normative_stock / days

        actual_expected = None
        actual_error = None

        if indigenous is not None and imported is not None:
            actual_expected = indigenous + imported
            actual_error = self._error_pct(
                actual_total,
                actual_expected,
            )

        return {
            "days": days,
            "daily": daily,
            "normative_stock": normative_stock,
            "actual_total": actual_total,
            "indigenous": indigenous,
            "imported": imported,
            "actual_expected": actual_expected,
            "actual_error": actual_error,
            "days_method": "explicit",
            "daily_method": "derived_from_normative_stock_div_days",
            "normative_days_col": self._excel_col(c["days"]),
            "daily_col": None,
            "normative_stock_col": self._excel_col(c["normative_stock"]),
            "indigenous_col": self._excel_col(c["indigenous"]),
            "import_col": self._excel_col(c["import"]),
            "actual_total_col": self._excel_col(c["actual_total"]),
        }

    def _extract_xlsx_standard(
        self,
        values: list[list[Any]],
        row: int,
    ) -> dict[str, Any]:
        c = self.XLSX_STANDARD

        daily = self._num(self._get(values, row, c["daily"]))
        normative_stock = self._num(
            self._get(values, row, c["normative_stock"])
        )
        indigenous = self._num(
            self._get(values, row, c["indigenous"])
        )
        imported = self._num(
            self._get(values, row, c["import"])
        )
        actual_total = self._num(
            self._get(values, row, c["actual_total"])
        )

        if daily is None or daily <= 0:
            raise RuntimeError(
                f"Invalid XLSX Daily Requirement: {daily}"
            )

        if normative_stock is None or normative_stock <= 0:
            raise RuntimeError(
                f"Invalid XLSX Normative Stock: {normative_stock}"
            )

        if actual_total is None or actual_total < 0:
            raise RuntimeError(
                f"Invalid XLSX Actual Total: {actual_total}"
            )

        # XLSX_STANDARD does not expose Normative Days directly.
        normative_days = normative_stock / daily

        actual_expected = None
        actual_error = None

        if indigenous is not None and imported is not None:
            actual_expected = indigenous + imported
            actual_error = self._error_pct(
                actual_total,
                actual_expected,
            )

        return {
            "days": normative_days,
            "daily": daily,
            "normative_stock": normative_stock,
            "actual_total": actual_total,
            "indigenous": indigenous,
            "imported": imported,
            "actual_expected": actual_expected,
            "actual_error": actual_error,
            "days_method": "derived_from_normative_stock_div_daily",
            "daily_method": "report_cell",
            "normative_days_col": None,
            "daily_col": self._excel_col(c["daily"]),
            "normative_stock_col": self._excel_col(c["normative_stock"]),
            "indigenous_col": self._excel_col(c["indigenous"]),
            "import_col": self._excel_col(c["import"]),
            "actual_total_col": self._excel_col(c["actual_total"]),
        }

    def extract_file(self, path: Path) -> ExtractionResult:
        report_date = self._date_from_filename(path.name)

        sheet_name, values = self.load_workbook(path)

        grand_total_row = self.find_grand_total_row(values)

        layout = self.detect_layout(
            path,
            values,
            grand_total_row,
        )

        if layout == "XLS_OLD":
            data = self._extract_xls_old(
                values,
                grand_total_row,
            )
        elif layout == "XLS_NEW":
            data = self._extract_xls_new(
                values,
                grand_total_row,
            )
        elif layout == "XLSX_STANDARD":
            data = self._extract_xlsx_standard(
                values,
                grand_total_row,
            )
        else:
            raise RuntimeError(f"Unsupported layout: {layout}")

        # --------------------------------------------------------
        # Validation 1: Daily × Days = Normative Stock
        # --------------------------------------------------------
        reconstructed_stock = (
            data["daily"] * data["days"]
        )

        normative_error = self._error_pct(
            reconstructed_stock,
            data["normative_stock"],
        )

        if normative_error > 3.0:
            raise RuntimeError(
                "Normative validation failed: "
                f"Daily={data['daily']}, "
                f"Days={data['days']}, "
                f"Normative Stock={data['normative_stock']}, "
                f"Calculated={reconstructed_stock}, "
                f"Error={normative_error:.4f}%"
            )

        # --------------------------------------------------------
        # Validation 2: Indigenous + Import ≈ Total
        # --------------------------------------------------------
        if data["actual_error"] is not None:
            if data["actual_error"] > 3.0:
                raise RuntimeError(
                    "Actual Stock validation failed: "
                    f"Expected={data['actual_expected']}, "
                    f"Actual={data['actual_total']}, "
                    f"Error={data['actual_error']:.4f}%"
                )

        return ExtractionResult(
            report_date=report_date,
            file_name=path.name,
            sheet_name=sheet_name,
            layout=layout,
            grand_total_row=grand_total_row + 1,

            normative_days=self._clean_num(data["days"]),
            daily_requirement=self._clean_num(data["daily"]),
            normative_stock=self._clean_num(data["normative_stock"]),
            actual_stock_total=self._clean_num(data["actual_total"]),

            normative_days_column=data["normative_days_col"],
            daily_column=data["daily_col"],
            normative_stock_column=data["normative_stock_col"],
            indigenous_column=data["indigenous_col"],
            import_column=data["import_col"],
            actual_total_column=data["actual_total_col"],

            days_method=data["days_method"],
            daily_method=data["daily_method"],

            normative_validation_error_pct=round(
                normative_error,
                6,
            ),
            actual_validation_error_pct=(
                round(data["actual_error"], 6)
                if data["actual_error"] is not None
                else None
            ),
        )

    # ============================================================
    # CSV STORAGE
    # ============================================================

    def _read_output(self) -> pd.DataFrame:
        if not self.output_csv.exists():
            return pd.DataFrame(columns=self.OUTPUT_COLUMNS)

        try:
            df = pd.read_csv(self.output_csv)
        except Exception:
            return pd.DataFrame(columns=self.OUTPUT_COLUMNS)

        for col in self.OUTPUT_COLUMNS:
            if col not in df.columns:
                df[col] = pd.NA

        return df[self.OUTPUT_COLUMNS]

    def _read_audit(self) -> pd.DataFrame:
        if not self.audit_csv.exists():
            return pd.DataFrame()

        try:
            return pd.read_csv(self.audit_csv)
        except Exception:
            return pd.DataFrame()

    def save_result(self, result: ExtractionResult) -> None:
        """
        Upsert one date.

        This is important for daily execution:
        if today's/ yesterday's report is rerun, the CSV does not receive
        duplicate rows.
        """
        output_df = self._read_output()

        new_row = pd.DataFrame(
            [result.output_dict()],
            columns=self.OUTPUT_COLUMNS,
        )

        output_df = pd.concat(
            [output_df, new_row],
            ignore_index=True,
        )

        output_df["Date"] = pd.to_datetime(
            output_df["Date"],
            errors="coerce",
        )

        output_df = (
            output_df
            .dropna(subset=["Date"])
            .sort_values("Date")
            .drop_duplicates(
                subset=["Date"],
                keep="last",
            )
            .reset_index(drop=True)
        )

        output_df["Date"] = output_df["Date"].dt.strftime("%Y-%m-%d")

        output_df.to_csv(
            self.output_csv,
            index=False,
        )

        # Audit is also an upsert by Date.
        audit_df = self._read_audit()
        audit_row = pd.DataFrame([result.audit_dict()])

        if audit_df.empty:
            audit_df = audit_row
        else:
            audit_df = pd.concat(
                [audit_df, audit_row],
                ignore_index=True,
            )

            audit_df["_DateKey"] = pd.to_datetime(
                audit_df["Date"],
                errors="coerce",
            )

            audit_df = (
                audit_df
                .dropna(subset=["_DateKey"])
                .sort_values("_DateKey")
                .drop_duplicates(
                    subset=["Date"],
                    keep="last",
                )
                .drop(columns=["_DateKey"])
                .reset_index(drop=True)
            )

        audit_df.to_csv(
            self.audit_csv,
            index=False,
        )

    def _save_failure(
        self,
        report_date: str,
        error: Exception,
    ) -> None:
        row = pd.DataFrame(
            [
                {
                    "Date": report_date,
                    "Error": str(error),
                    "Timestamp": datetime.now().isoformat(
                        timespec="seconds"
                    ),
                }
            ]
        )

        if self.failed_csv.exists():
            try:
                old = pd.read_csv(self.failed_csv)
                row = pd.concat(
                    [old, row],
                    ignore_index=True,
                )
            except Exception:
                pass

        row.to_csv(
            self.failed_csv,
            index=False,
        )

    # ============================================================
    # SINGLE-DAY PIPELINE
    # ============================================================

    def process_date(
        self,
        report_date: date | datetime | str,
    ) -> Optional[ExtractionResult]:
        """
        Download -> extract -> validate -> store for exactly one date.
        """
        if isinstance(report_date, datetime):
            report_date = report_date.date()
        elif isinstance(report_date, str):
            report_date = datetime.strptime(
                report_date,
                "%Y-%m-%d",
            ).date()

        date_string = report_date.strftime("%Y-%m-%d")

        try:
            path = self.download_report(report_date)

            if path is None:
                raise RuntimeError(
                    f"NPP report unavailable for {date_string}"
                )

            print(f"  Reading: {path.name}")

            result = self.extract_file(path)

            self.save_result(result)

            print(
                f"  ✓ Extracted | "
                f"Days={result.normative_days} | "
                f"Daily={result.daily_requirement} | "
                f"Actual Total={result.actual_stock_total}"
            )

            print(
                f"    Layout={result.layout}, "
                f"Grand Total Row={result.grand_total_row}, "
                f"Normative Error="
                f"{result.normative_validation_error_pct}%"
            )

            if result.actual_validation_error_pct is not None:
                print(
                    f"    Actual Error="
                    f"{result.actual_validation_error_pct}%"
                )

            return result

        except Exception as exc:
            print(f"  ✗ FAILED {date_string}: {exc}")
            self._save_failure(date_string, exc)
            return None

    # ============================================================
    # DAILY / BACKFILL
    # ============================================================

    def run_daily(
        self,
        report_date: date | datetime | str | None = None,
    ) -> Optional[ExtractionResult]:
        """
        Daily entry point.

        With no argument:
            automatically processes yesterday.

        This means the same code can be scheduled every day without
        changing the date in the code.
        """
        if report_date is None:
            report_date = datetime.now().date() - timedelta(days=1)

        result = self.process_date(report_date)

        print("\n" + "=" * 80)
        if result:
            print(
                f"DAILY RUN SUCCESSFUL: {result.report_date}"
            )
            print(f"CSV:   {self.output_csv}")
            print(f"AUDIT: {self.audit_csv}")
        else:
            print(f"DAILY RUN FAILED: {report_date}")
        print("=" * 80)

        return result

    def run_backfill(
        self,
        start_date: date | datetime | str,
        end_date: date | datetime | str,
    ) -> list[ExtractionResult]:
        """
        Backfill an inclusive date range.

        Useful once, when building the historical dataset. After that,
        run_daily() is sufficient.
        """
        def as_date(value):
            if isinstance(value, datetime):
                return value.date()
            if isinstance(value, date):
                return value
            return datetime.strptime(
                value,
                "%Y-%m-%d",
            ).date()

        start = as_date(start_date)
        end = as_date(end_date)

        if start > end:
            raise ValueError(
                f"start_date {start} is after end_date {end}"
            )

        results: list[ExtractionResult] = []
        failures = []

        current = start

        while current <= end:
            result = self.process_date(current)

            if result is not None:
                results.append(result)
            else:
                failures.append(current.isoformat())

            current += timedelta(days=1)

        print("\n" + "=" * 80)
        print("BACKFILL SUMMARY")
        print("=" * 80)
        print(f"Requested : {(end - start).days + 1}")
        print(f"Success   : {len(results)}")
        print(f"Failed    : {len(failures)}")

        if failures:
            print("Failed dates:")
            for d in failures:
                print(f"  - {d}")

        print(f"\nCSV:   {self.output_csv}")
        print(f"AUDIT: {self.audit_csv}")
        print("=" * 80)

        return results

    # ============================================================
    # OPTIONAL SELF-CHECK / VIEW
    # ============================================================

    def get_data(self) -> pd.DataFrame:
        """Return the current final CSV as a DataFrame."""
        return self._read_output()

    def print_latest(self, rows: int = 10) -> None:
        df = self._read_output()

        if df.empty:
            print("No extracted data yet.")
            return

        print(
            df.tail(rows).to_string(index=False)
        )


# ================================================================
# DIRECT EXECUTION
# ================================================================
#
# The class itself contains the entire pipeline.
#
# Running:
#     python npp_coal_stock_pipeline.py
#
# automatically processes yesterday.
#
# For historical loading, import the class and call run_backfill().
# ================================================================

if __name__ == "__main__":
    pipeline = NPPCoalStockPipeline()
    pipeline.run_backfill(
        start_date=date(2026, 8, 12),
        end_date=date(2026, 9, 8)
    )
    pipeline.run_daily()
