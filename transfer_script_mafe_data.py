import csv
import re
import sys
import datetime
from pathlib import Path
import win32com.client
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

JOBS = [
    {
        "name": "LEC",
        "source_file": r"C:\MAF_Reports_2026\data_sourse\MAFLEC.xlsx",
        "skip_top_rows": 10,
        "destination_file": r"C:\MAF_Reports_2026\data_sheets\MAFE Business Data.xlsx",
        "sheet": "Tickets Raw Data",
        "table": "Tickets_RawData",
        "start_row": 85597,
        "last_column": "TimeTaken",
        "column_map": {"StartTime": "Start Date"},
        "clear_source": False,
    },
    {
        "name": "FCR",
        "source_file": r"C:\MAF_Reports_2026\data_sourse\FCR.xlsx",
        "skip_top_rows": 14,
        "destination_file": r"C:\MAF_Reports_2026\data_sheets\MAFE Business Data.xlsx",
        "sheet": "FCR ",
        "table": "FCR_Rawdata",
        "start_row": 103858,
        "last_column": None,
        "column_map": {},
        "clear_source": False,
    },
    {
        "name": "NPS",
        "source_file": r"C:\MAF_Reports_2026\data_sourse\NPSE.csv",
        "skip_top_rows": 0,
        "destination_file": r"C:\MAF_Reports_2026\data_sheets\MAFE Business Data.xlsx",
        "sheet": "NPS RawData",
        "table": "NPS_RawData",
        "start_row": 1558,
        "last_column": None,
        "column_map": {},
        "clear_source": False,
    },
]

STRICT_HEADER_MATCH = True
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
DATETIME_PATTERN = re.compile(
    r"^\d{1,2}/\d{1,2}/\d{4} \d{1,2}:\d{2}:\d{2} [AP]M$"
)
# أعمدة الجدول اللي اسمها Column1 / Column2 ... بتتتجاهل (مش مطلوب تيجي من السورس)
GENERIC_COLUMN_PATTERN = re.compile(r"^Column\d+$", re.IGNORECASE)
# أعمدة السورس اللي ملهاش اسم بتتتجاهل بصمت (True = اطبع رسالة عنها)
SHOW_UNNAMED_COLUMN_MESSAGES = False


class TransferError(Exception):
    pass


def norm_header(value):
    return re.sub(r"\s+", " ", str(value)).strip().casefold()


def make_unique_headers(headers):
    result = []
    used = set()
    for header in headers:
        base = str(header).strip()
        candidate = base
        counter = 2
        while norm_header(candidate) in used:
            candidate = f"{base}{counter}"
            counter += 1
        result.append(candidate)
        used.add(norm_header(candidate))
    return result


def is_blank(value):
    return value is None or (isinstance(value, str) and value.strip() == "")


def to_com_value(value):
    if is_blank(value):
        return None
    if isinstance(value, datetime.datetime):
        return value
    if isinstance(value, datetime.date):
        return datetime.datetime(value.year, value.month, value.day)
    if isinstance(value, datetime.time):
        return (
            value.hour * 3600
            + value.minute * 60
            + value.second
            + value.microsecond / 1e6
        ) / 86400
    if isinstance(value, datetime.timedelta):
        return value.total_seconds() / 86400
    if isinstance(value, str):
        text = value.strip()
        if DATETIME_PATTERN.match(text):
            try:
                return datetime.datetime.strptime(
                    text, "%m/%d/%Y %I:%M:%S %p"
                )
            except ValueError:
                pass
        if value.startswith("="):
            return "'" + value
    return value


def read_values(rng):
    value = rng.Value
    if not isinstance(value, tuple):
        return [[value]]
    return [list(row) for row in value]


# ------------------------------------------------------------
# قراءة السورس
# ------------------------------------------------------------
def read_source_sheet(job):
    path = Path(job["source_file"])
    skip = job.get("skip_top_rows", 0)
    last_column = job.get("last_column")
    if not path.is_file():
        raise TransferError(f"Source file not found: {path}")

    if path.suffix.lower() == ".csv":
        rows = None
        last_error = None
        for encoding in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                with path.open("r", newline="", encoding=encoding) as file:
                    sample = file.read(8192)
                    file.seek(0)
                    try:
                        dialect = csv.Sniffer().sniff(
                            sample, delimiters=",;\t|"
                        )
                    except csv.Error:
                        dialect = csv.excel
                    rows = [list(r) for r in csv.reader(file, dialect)]
                break
            except UnicodeError as exc:
                last_error = exc
        if rows is None:
            raise TransferError(
                f"Could not read CSV {path.name}: {last_error}"
            )
        label = path.name
    else:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            if not wb.sheetnames:
                raise TransferError(f"{path.name} has no worksheets.")
            ws = wb[wb.sheetnames[0]]
            label = f"{path.name}/{ws.title}"
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
        finally:
            wb.close()

    if skip:
        rows = rows[skip:]
    rows = [r for r in rows if any(not is_blank(v) for v in r)]
    if not rows:
        raise TransferError(f"Source {label} is empty after skipping rows.")

    raw_headers = rows[0]
    data_rows = rows[1:]

    keep = []
    for i, header in enumerate(raw_headers):
        if is_blank(header):
            if SHOW_UNNAMED_COLUMN_MESSAGES and any(
                i < len(r) and not is_blank(r[i]) for r in data_rows
            ):
                print(
                    f"  EMPTY HEADER SKIPPED | Column {get_column_letter(i + 1)}"
                )
        else:
            keep.append(i)

    if last_column:
        stop = None
        for pos, i in enumerate(keep):
            if norm_header(raw_headers[i]) == norm_header(last_column):
                stop = pos
        if stop is None:
            print(
                f"  WARNING | last_column '{last_column}' not found in "
                f"source, using all named columns."
            )
        else:
            dropped = [str(raw_headers[i]).strip() for i in keep[stop + 1:]]
            if dropped:
                print(f"  COLUMNS EXCLUDED | {', '.join(dropped)}")
            keep = keep[:stop + 1]

    cmap = {norm_header(k): v for k, v in job.get("column_map", {}).items()}
    original = []
    for i in keep:
        header = str(raw_headers[i]).strip()
        new = cmap.get(norm_header(header))
        if new is not None:
            print(f"  COLUMN RENAMED | {header} -> {new}")
            header = new
        original.append(header)
    headers = make_unique_headers(original)

    values = [
        [row[i] if i < len(row) else None for i in keep]
        for row in data_rows
    ]
    return headers, values


# ------------------------------------------------------------
# الكتابة في الجدول (استبدال من صف البداية، مش إضافة)
# ------------------------------------------------------------
def find_worksheet(wb, sheet_name):
    names = []
    for sheet in wb.Worksheets:
        names.append(sheet.Name)
        if sheet.Name == sheet_name:
            return sheet
    raise TransferError(
        f"Sheet '{sheet_name}' not found. Available: {', '.join(names)}"
    )


def find_table(ws, table_name):
    names = []
    for table in ws.ListObjects:
        names.append(table.Name)
        if table.Name == table_name:
            return table
    raise TransferError(
        f"Table '{table_name}' not found on '{ws.Name}'. "
        f"Available: {', '.join(names)}"
    )


def replace_table_data(wb, job, source_headers, source_rows):
    ws = find_worksheet(wb, job["sheet"])
    table = find_table(ws, job["table"])
    if table.ShowTotals:
        raise TransferError(
            f"Table '{table.Name}' has a Totals row. Turn it off first."
        )
    if not source_rows:
        raise TransferError(
            "Source has headers but no data rows. Nothing was changed."
        )

    header_row = table.HeaderRowRange.Row
    min_col = table.Range.Column
    first_data_row = header_row + 1
    old_last = table.Range.Row + table.Range.Rows.Count - 1
    dest_headers = [c.Name for c in table.ListColumns]
    col_count = len(dest_headers)
    max_col = min_col + col_count - 1

    start_row = job.get("start_row") or first_data_row
    if start_row < first_data_row:
        raise TransferError(
            f"start_row ({start_row}) is above the first data row "
            f"of the table ({first_data_row})."
        )

    # ربط أعمدة السورس بأعمدة الجدول
    dest_index = {}
    for i, h in enumerate(make_unique_headers(dest_headers)):
        dest_index[norm_header(h)] = i
    src_to_dest = {}
    unknown = []
    for si, h in enumerate(source_headers):
        di = dest_index.get(norm_header(h))
        if di is None:
            unknown.append(h)
        else:
            src_to_dest[si] = di
    if unknown:
        if job.get("skip_unknown_source_columns", True):
            print(
                f"  WARNING | {len(unknown)} source column(s) not in table "
                f"'{table.Name}' were SKIPPED: {', '.join(unknown)}"
            )
        else:
            raise TransferError(
                f"Source headers not found in table '{table.Name}': "
                f"{', '.join(unknown)}"
            )
    mapped = set(src_to_dest.values())

    protected_start = col_count
    if job.get("last_column"):
        li = dest_index.get(norm_header(job["last_column"]))
        if li is not None:
            protected_start = li + 1

    def is_formula_column(i):
        if old_last < first_data_row:
            return False
        return bool(ws.Cells(first_data_row, min_col + i).HasFormula)

    missing = [
        dest_headers[i]
        for i in range(col_count)
        if i not in mapped
        and i < protected_start
        and not GENERIC_COLUMN_PATTERN.match(str(dest_headers[i]).strip())
        and not is_formula_column(i)
    ]
    if missing and STRICT_HEADER_MATCH:
        raise TransferError(
            f"Table columns missing from source: {', '.join(missing)}"
        )

    new_rows = []
    for src in source_rows:
        out = [None] * col_count
        for si, di in src_to_dest.items():
            out[di] = src[si]
        new_rows.append(out)

    end_row = start_row + len(new_rows) - 1

    # لو الجدول هيكبر: تأكد إن اللي تحته فاضي
    if end_row > old_last:
        below = ws.Range(
            ws.Cells(old_last + 1, min_col), ws.Cells(end_row, max_col)
        )
        if wb.Application.WorksheetFunction.CountA(below) > 0:
            raise TransferError(
                f"Cannot expand '{table.Name}': cells below it contain data."
            )

    # خد الفورمات من أول صف داتا قبل ما نغير حاجة
    formats = {}
    if old_last >= start_row:
        for di in mapped:
            formats[di] = ws.Cells(start_row, min_col + di).NumberFormat

    # لو الداتا الجديدة أقصر من القديمة: امسح الصفوف الزيادة
    if end_row < old_last:
        ws.Range(
            ws.Cells(end_row + 1, min_col), ws.Cells(old_last, max_col)
        ).ClearContents()

    # غيّر حجم الجدول ليطابق الداتا الجديدة
    if end_row != old_last:
        table.Resize(
            ws.Range(ws.Cells(header_row, min_col), ws.Cells(end_row, max_col))
        )

    # اكتب الداتا
    for di in mapped:
        col = min_col + di
        target = ws.Range(ws.Cells(start_row, col), ws.Cells(end_row, col))
        if di in formats:
            target.NumberFormat = formats[di]
        target.Value = tuple((to_com_value(r[di]),) for r in new_rows)

    # املا أعمدة المعادلات (اللي مش جاية من السورس) لحد آخر صف
    template_row = min(old_last, end_row)
    if template_row >= first_data_row:
        for di in range(col_count):
            if di in mapped:
                continue
            col = min_col + di
            if (
                ws.Cells(template_row, col).HasFormula
                and not ws.Cells(end_row, col).HasFormula
            ):
                ws.Range(
                    ws.Cells(template_row, col), ws.Cells(end_row, col)
                ).FillDown()

    return len(new_rows), start_row, end_row


# ------------------------------------------------------------
# مسح السورس (اختياري)
# ------------------------------------------------------------
def clear_source_data(job):
    path = Path(job["source_file"])
    skip = job.get("skip_top_rows", 0)
    header_line = skip + 1
    if path.suffix.lower() == ".csv":
        with path.open("r", newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
        with path.open("w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows(rows[:header_line])
        return
    wb = load_workbook(path)
    try:
        ws = wb[wb.sheetnames[0]]
        if ws.max_row > header_line:
            for row in ws.iter_rows(
                min_row=header_line + 1,
                max_row=ws.max_row,
                min_col=1,
                max_col=ws.max_column,
            ):
                for cell in row:
                    cell.value = None
        wb.save(path)
    finally:
        wb.close()


# ------------------------------------------------------------
# التشغيل
# ------------------------------------------------------------
def run_job(excel, job):
    print(f"\n=== {job['name']} ===")
    headers, rows = read_source_sheet(job)
    print(f"READ OK | {Path(job['source_file']).name} | {len(rows)} rows")

    dest_path = Path(job["destination_file"]).resolve()
    if not dest_path.is_file():
        raise TransferError(f"Destination not found: {dest_path}")

    workbook = None
    try:
        workbook = excel.Workbooks.Open(
            str(dest_path), UpdateLinks=0, ReadOnly=False
        )
        excel.Calculation = XL_CALCULATION_MANUAL
        count, start_row, end_row = replace_table_data(
            workbook, job, headers, rows
        )
        excel.Calculation = XL_CALCULATION_AUTOMATIC
        workbook.Save()
        print(
            f"TRANSFER OK | {job['sheet']}/{job['table']} | "
            f"rows={count}, from row {start_row} to {end_row}"
        )
    finally:
        try:
            excel.Calculation = XL_CALCULATION_AUTOMATIC
        except Exception:
            pass
        if workbook is not None:
            workbook.Close(SaveChanges=False)

    if job.get("clear_source"):
        try:
            clear_source_data(job)
            print("CLEARED | source data removed (header kept)")
        except Exception as exc:
            print(f"CLEAR FAILED | {exc}")


def main():
    excel = None
    failed = []
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        for job in JOBS:
            try:
                run_job(excel, job)
            except Exception as exc:
                failed.append(job["name"])
                print(f"FAILED | {job['name']} | {exc}")
                print("  Destination for this job was NOT changed.")
    finally:
        if excel is not None:
            excel.Quit()

    print()
    if failed:
        print(f"STATUS: FINISHED WITH ERRORS in: {', '.join(failed)}")
        return 1
    print("STATUS: SUCCESS")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as error:
        print(f"\nUNEXPECTED ERROR: {type(error).__name__}: {error}")
        sys.exit(1)