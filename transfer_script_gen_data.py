import csv
import re
import sys
import datetime
from pathlib import Path
import win32com.client
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

TRANSFER_JOBS = [
    {
        "source_file": r"C:\MAF_Reports_2026\data_sourse\d.csv",
        "destination_file": r"C:\MAF_Reports_2026\data_sheets\Genyses Raw Data..xlsx",
        "destination_sheet": "DBU Queues",
        "destination_table": "DBU_Queues",
    },
    {
        "source_file": r"C:\MAF_Reports_2026\data_sourse\e.csv",
        "destination_file": r"C:\MAF_Reports_2026\data_sheets\Genyses Raw Data..xlsx",
        "destination_sheet": "MAFE Queues",
        "destination_table": "MAFE_Queues",
    },
    {
        "source_file": r"C:\MAF_Reports_2026\data_sourse\f.csv",
        "destination_file": r"C:\MAF_Reports_2026\data_sheets\Genyses Raw Data..xlsx",
        "destination_sheet": "Queue Performance-Fintech",
        "destination_table": "Fintech_Queue",
    },
    {
        "source_file": r"C:\MAF_Reports_2026\data_sourse\ff.xlsx",
        "destination_file": r"C:\MAF_Reports_2026\data_sheets\Genyses Raw Data..xlsx",
        "destination_sheet": "General CSAT",
        "destination_table": "General_CSAT",
        "skip_top_rows": 7,
    },
    {
        "source_file": r"C:\MAF_Reports_2026\data_sourse\ee.xlsx",
        "destination_file": r"C:\MAF_Reports_2026\data_sheets\Genyses Raw Data..xlsx",
        "destination_sheet": "General CSAT",
        "destination_table": "General_CSAT",
        "skip_top_rows": 7,
    },
    {
        "source_file": r"C:\MAF_Reports_2026\data_sourse\Skills Performance.csv",
        "destination_file": r"C:\MAF_Reports_2026\data_sheets\Genyses Raw Data..xlsx",
        "destination_sheet": "Skill Performance",
        "destination_table": "Skill_RawData",
    },
    {
        "source_file": r"C:\MAF_Reports_2026\data_sourse\Agent Performance Summary.csv",
        "destination_file": r"C:\MAF_Reports_2026\data_sheets\Genyses Raw Data..xlsx",
        "destination_sheet": "Agent Performance_Raw Data",
        "destination_table": "Agent_Performance",
    },
    {
        "source_file": r"C:\MAF_Reports_2026\data_sourse\Agent Status Summary.csv",
        "destination_file": r"C:\MAF_Reports_2026\data_sheets\Genyses Raw Data..xlsx",
        "destination_sheet": "Agent Status Summary",
        "destination_table": "Status_Summary",
    },
]
CLEAR_SOURCE_DATA = True
STRICT_HEADER_MATCH = True
DUPLICATE_KEY_COLUMNS = []
SKIP_JOBS = []
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105


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
    return value is None or (
        isinstance(value, str) and value.strip() == ""
    )


def norm_value(value):
    if is_blank(value):
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = re.sub(r"\s+", " ", value).strip().casefold()
        try:
            return float(text)
        except ValueError:
            return text
    return value


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
    if isinstance(value, str) and value.startswith("="):
        return "'" + value
    return value


def read_values(rng):
    value = rng.Value
    if not isinstance(value, tuple):
        return [[value]]
    return [list(row) for row in value]


def read_source_sheet(
    path,
    skip_top_rows=0,
    last_column=None,
    ignore_columns=None
):
    path = Path(path)
    if not path.is_file():
        raise TransferError(
            f"Source file not found: {path}"
        )
    if path.suffix.lower() == ".csv":
        rows = None
        last_error = None
        for encoding in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                with path.open(
                    "r",
                    newline="",
                    encoding=encoding
                ) as file:
                    sample = file.read(8192)
                    file.seek(0)
                    try:
                        dialect = csv.Sniffer().sniff(
                            sample,
                            delimiters=",;\t|"
                        )
                    except csv.Error:
                        dialect = csv.excel
                    rows = [
                        list(row)
                        for row in csv.reader(file, dialect)
                    ]
                break
            except UnicodeError as exc:
                last_error = exc
        if rows is None:
            raise TransferError(
                f"Could not read CSV file {path.name}: {last_error}"
            )
        source_label = path.name
    else:
        keep_vba = path.suffix.lower() == ".xlsm"
        wb = load_workbook(
            path,
            read_only=True,
            data_only=True,
            keep_vba=keep_vba
        )
        try:
            if not wb.sheetnames:
                raise TransferError(
                    f"Source workbook {path.name} has no worksheets."
                )
            ws = wb[wb.sheetnames[0]]
            source_label = f"{path.name}/{ws.title}"
            rows = [
                list(row)
                for row in ws.iter_rows(values_only=True)
            ]
        finally:
            wb.close()
    if skip_top_rows:
        rows = rows[skip_top_rows:]
    rows = [
        row for row in rows
        if any(not is_blank(v) for v in row)
    ]
    if not rows:
        raise TransferError(
            f"Source {source_label} is empty."
        )
    raw_headers = rows[0]
    data_rows = rows[1:]
    keep_columns = []
    for i, header in enumerate(raw_headers):
        if is_blank(header):
            if any(
                i < len(row) and not is_blank(row[i])
                for row in data_rows
            ):
                print(
                    f"EMPTY HEADER SKIPPED | {source_label} | "
                    f"Column {get_column_letter(i + 1)}"
                )
        else:
            keep_columns.append(i)
    if last_column is not None:
        stop = None
        for position, i in enumerate(keep_columns):
            if norm_header(raw_headers[i]) == norm_header(last_column):
                stop = position
        if stop is None:
            raise TransferError(
                f"Column '{last_column}' not found in {source_label}. "
                f"Headers found: "
                f"{', '.join(str(raw_headers[i]).strip() for i in keep_columns)}"
            )
        dropped = [
            str(raw_headers[i]).strip()
            for i in keep_columns[stop + 1:]
        ]
        if dropped:
            print(
                f"COLUMNS EXCLUDED | {source_label} | "
                f"{', '.join(dropped)}"
            )
        keep_columns = keep_columns[:stop + 1]
    if ignore_columns:
        ignored = {norm_header(name) for name in ignore_columns}
        skipped_names = [
            str(raw_headers[i]).strip()
            for i in keep_columns
            if norm_header(raw_headers[i]) in ignored
        ]
        if skipped_names:
            print(
                f"COLUMNS IGNORED | {source_label} | "
                f"{', '.join(skipped_names)}"
            )
        keep_columns = [
            i for i in keep_columns
            if norm_header(raw_headers[i]) not in ignored
        ]
    original_headers = [
        str(raw_headers[i]).strip()
        for i in keep_columns
    ]
    headers = make_unique_headers(original_headers)
    if headers != original_headers:
        print(
            f"DUPLICATE HEADERS FIXED | {source_label}"
        )
        for old, new in zip(original_headers, headers):
            if old != new:
                print(f"  {old} -> {new}")
    values = [
        [
            row[i] if i < len(row) else None
            for i in keep_columns
        ]
        for row in data_rows
    ]
    return headers, values


def find_worksheet(wb, sheet_name, destination_file):
    names = []
    for sheet in wb.Worksheets:
        names.append(sheet.Name)
        if sheet.Name == sheet_name:
            return sheet
    raise TransferError(
        f"Destination sheet '{sheet_name}' not found in "
        f"{destination_file}. Available: {', '.join(names)}"
    )


def find_table(ws, table_name):
    names = []
    for table in ws.ListObjects:
        names.append(table.Name)
        if table.Name == table_name:
            return table
    raise TransferError(
        f"Table '{table_name}' not found on destination sheet "
        f"'{ws.Name}'. Available tables: {', '.join(names)}"
    )


def append_to_table(wb, job, source_headers, source_rows):
    ws = find_worksheet(
        wb,
        job["destination_sheet"],
        job["destination_file"]
    )
    table = find_table(
        ws,
        job["destination_table"]
    )
    if table.ShowTotals:
        raise TransferError(
            f"Table '{table.Name}' has a Totals row. "
            f"Turn it off before running."
        )
    header_row = table.HeaderRowRange.Row
    min_col = table.Range.Column
    first_data_row = header_row + 1
    last_row = table.Range.Row + table.Range.Rows.Count - 1
    destination_headers = [
        column.Name
        for column in table.ListColumns
    ]
    column_count = len(destination_headers)
    max_col = min_col + column_count - 1
    unique_destination_headers = make_unique_headers(
        destination_headers
    )
    destination_index = {}
    for i, header in enumerate(unique_destination_headers):
        key = norm_header(header)
        if key in destination_index:
            raise TransferError(
                f"Duplicate destination header mapping: '{header}'."
            )
        destination_index[key] = i
    source_to_destination = {}
    unknown = []
    for source_index, header in enumerate(source_headers):
        destination_index_for_header = destination_index.get(
            norm_header(header)
        )
        if destination_index_for_header is None:
            unknown.append(header)
        else:
            source_to_destination[source_index] = (
                destination_index_for_header
            )
    if unknown:
        raise TransferError(
            f"Source headers not found in table '{table.Name}': "
            f"{', '.join(unknown)}. Table columns are: "
            f"{', '.join(destination_headers)}"
        )
    mapped_columns = set(
        source_to_destination.values()
    )
    protected_start = column_count
    if job.get("last_column"):
        last_index = destination_index.get(
            norm_header(job["last_column"])
        )
        if last_index is None:
            raise TransferError(
                f"Column '{job['last_column']}' not found in "
                f"table '{table.Name}'."
            )
        protected_start = last_index + 1
    missing = [
        unique_destination_headers[i]
        for i in range(column_count)
        if i not in mapped_columns and i < protected_start
    ]
    if missing and STRICT_HEADER_MATCH:
        raise TransferError(
            f"Destination columns missing from source for table "
            f"'{table.Name}': {', '.join(missing)}"
        )
    new_rows = []
    for source_row in source_rows:
        output_row = [None] * column_count
        for source_i, destination_i in source_to_destination.items():
            output_row[destination_i] = source_row[source_i]
        new_rows.append(output_row)
    skipped = 0
    if DUPLICATE_KEY_COLUMNS:
        key_indexes = []
        for key_name in DUPLICATE_KEY_COLUMNS:
            idx = destination_index.get(
                norm_header(key_name)
            )
            if idx is None or idx not in mapped_columns:
                raise TransferError(
                    f"Duplicate key '{key_name}' must exist "
                    f"in source and destination."
                )
            key_indexes.append(idx)

        def make_key(row):
            return tuple(
                norm_value(row[i])
                for i in key_indexes
            )

        existing_keys = set()
        if last_row >= first_data_row:
            existing_rows = read_values(
                ws.Range(
                    ws.Cells(first_data_row, min_col),
                    ws.Cells(last_row, max_col)
                )
            )
            for row in existing_rows:
                if any(not is_blank(v) for v in row):
                    existing_keys.add(make_key(row))
        unique_rows = []
        for row in new_rows:
            key = make_key(row)
            if key in existing_keys:
                skipped += 1
            else:
                existing_keys.add(key)
                unique_rows.append(row)
        new_rows = unique_rows
    if not new_rows:
        return 0, skipped
    if last_row < first_data_row:
        table_empty = True
    elif last_row == first_data_row:
        first_row_values = read_values(
            ws.Range(
                ws.Cells(first_data_row, min_col),
                ws.Cells(first_data_row, max_col)
            )
        )[0]
        table_empty = all(
            is_blank(first_row_values[i])
            for i in mapped_columns
        )
    else:
        table_empty = False
    start_row = (
        first_data_row
        if table_empty
        else last_row + 1
    )
    end_row = start_row + len(new_rows) - 1
    new_last_row = max(last_row, end_row)
    if not table_empty and end_row > last_row:
        below = ws.Range(
            ws.Cells(last_row + 1, min_col),
            ws.Cells(end_row, max_col)
        )
        if wb.Application.WorksheetFunction.CountA(below) > 0:
            raise TransferError(
                f"Cannot expand table '{table.Name}': "
                f"cells below the table already contain data."
            )
    template_row = (
        first_data_row
        if table_empty
        else last_row
    )
    if new_last_row > last_row:
        table.Resize(
            ws.Range(
                ws.Cells(header_row, min_col),
                ws.Cells(new_last_row, max_col)
            )
        )
    for destination_i in range(column_count):
        column = min_col + destination_i
        target = ws.Range(
            ws.Cells(start_row, column),
            ws.Cells(end_row, column)
        )
        template = ws.Cells(template_row, column)
        if destination_i in mapped_columns:
            if not table_empty:
                target.NumberFormat = template.NumberFormat
            target.Value = tuple(
                (to_com_value(row[destination_i]),)
                for row in new_rows
            )
        elif template.HasFormula and not ws.Cells(end_row, column).HasFormula:
            ws.Range(
                ws.Cells(template_row, column),
                ws.Cells(end_row, column)
            ).FillDown()
    return len(new_rows), skipped


def transfer_to_workbook(dest_path, jobs, prepared_jobs):
    excel = None
    workbook = None
    results = []
    try:
        excel = win32com.client.DispatchEx(
            "Excel.Application"
        )
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        workbook = excel.Workbooks.Open(
            str(dest_path),
            UpdateLinks=0,
            ReadOnly=False
        )
        excel.Calculation = XL_CALCULATION_MANUAL
        for number, job in jobs:
            prepared = next(
                item
                for item in prepared_jobs
                if item[0] == number
            )
            _, _, headers, rows = prepared
            transferred, skipped = append_to_table(
                workbook,
                job,
                headers,
                rows
            )
            results.append((number, transferred, skipped))
            print(
                f"TRANSFER OK | Job {number}/8 | "
                f"{job['destination_sheet']}/"
                f"{job['destination_table']} | "
                f"added={transferred}, "
                f"duplicates_skipped={skipped}"
            )
        excel.Calculation = XL_CALCULATION_AUTOMATIC
        workbook.Save()
        return results
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()


def clear_source_data(job):
    path = Path(job["source_file"])

    if path.suffix.lower() == ".csv":
        with path.open(
            "r",
            newline="",
            encoding="utf-8-sig"
        ) as file:
            rows = list(csv.reader(file))

        if rows:
            with path.open(
                "w",
                newline="",
                encoding="utf-8-sig"
            ) as file:
                csv.writer(file).writerow(rows[0])

        return

    keep_vba = path.suffix.lower() == ".xlsm"

    wb = load_workbook(
        path,
        keep_vba=keep_vba
    )

    try:
        if not wb.sheetnames:
            raise TransferError(
                f"Cannot clear source: "
                f"{path.name} has no worksheets."
            )

        ws = wb[wb.sheetnames[0]]

        if ws.max_row >= 2:
            for row in ws.iter_rows(
                min_row=2,
                max_row=ws.max_row,
                min_col=1,
                max_col=ws.max_column
            ):
                for cell in row:
                    if not isinstance(cell, MergedCell):
                        cell.value = None

        wb.save(path)

    finally:
        wb.close()


def main():
    if len(TRANSFER_JOBS) != 8:
        print(
            f"CONFIGURATION ERROR: expected 8 jobs, "
            f"found {len(TRANSFER_JOBS)}."
        )
        return 1
    source_keys = set()
    destination_groups = {}
    for number, job in enumerate(
        TRANSFER_JOBS,
        start=1
    ):
        required = (
            "source_file",
            "destination_file",
            "destination_sheet",
            "destination_table"
        )
        missing = [
            key for key in required
            if not job.get(key)
        ]
        if missing:
            print(
                f"Job {number}: missing configuration: "
                f"{', '.join(missing)}"
            )
            return 1
        source_key = str(
            Path(job["source_file"]).resolve()
        ).casefold()
        if source_key in source_keys:
            print(
                f"CONFIGURATION ERROR: "
                f"duplicate source sheet in job {number}."
            )
            return 1
        source_keys.add(source_key)
        destination = str(
            Path(job["destination_file"]).resolve()
        )
        if number not in SKIP_JOBS:
            destination_groups.setdefault(
                destination,
                []
            ).append((number, job))
    prepared_jobs = []
    for number, job in enumerate(
        TRANSFER_JOBS,
        start=1
    ):
        if number in SKIP_JOBS:
            print(f"SKIPPED | Job {number}/8")
            continue
        try:
            headers, rows = read_source_sheet(
                job["source_file"],
                job.get("skip_top_rows", 0),
                job.get("last_column"),
                job.get("ignore_columns")
            )
            prepared_jobs.append(
                (number, job, headers, rows)
            )
            print(
                f"READ OK | Job {number}/8 | "
                f"{Path(job['source_file']).name} | "
                f"{len(rows)} rows"
            )
        except Exception as exc:
            print(
                f"READ FAILED | Job {number}/8 | {exc}"
            )
            print(
                "Stopped before modifying destination files."
            )
            return 1
    successful_destination_files = []
    totals = []
    for destination_path, jobs in destination_groups.items():
        dest_path = Path(destination_path)
        if not dest_path.is_file():
            print(
                f"DESTINATION NOT FOUND: {dest_path}"
            )
            return 1
        try:
            totals.extend(
                transfer_to_workbook(
                    dest_path,
                    jobs,
                    prepared_jobs
                )
            )
            successful_destination_files.append(
                str(dest_path)
            )
            print(f"SAVED: {dest_path}")
        except Exception as exc:
            print(
                f"TRANSFER/SAVE FAILED for "
                f"{dest_path}: {exc}"
            )
            print(
                "Source sheets have NOT been cleared."
            )
            return 1
    if CLEAR_SOURCE_DATA:
        print(
            "\nClearing source data rows "
            "(headers are preserved)..."
        )
        cleared = set()
        for number, job in enumerate(
            TRANSFER_JOBS,
            start=1
        ):
            key = str(
                Path(job["source_file"]).resolve()
            ).casefold()
            if key in cleared:
                continue
            try:
                clear_source_data(job)
                cleared.add(key)
                print(
                    f"CLEARED | Job {number}/8 | "
                    f"{Path(job['source_file']).name}"
                )
            except Exception as exc:
                print(
                    f"CLEAR FAILED | Job {number}/8 | "
                    f"{job['source_file']} | {exc}"
                )
                print(
                    "All destination transfers were completed, "
                    "but source cleanup was incomplete."
                )
                return 1
    print(
        "\n================ FINAL SUMMARY ================"
    )
    for number, transferred, skipped in sorted(totals):
        print(
            f"Job {number:02d}: "
            f"transferred={transferred}, "
            f"duplicates_skipped={skipped}"
        )
    print(
        f"Destination workbooks saved: "
        f"{len(successful_destination_files)}"
    )
    print(
        f"Source data cleared: "
        f"{'YES' if CLEAR_SOURCE_DATA else 'NO'}"
    )
    print("Backup created: NO")
    print("STATUS: SUCCESS")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as error:
        print(
            f"\nUNEXPECTED ERROR: "
            f"{type(error).__name__}: {error}"
        )
        print(
            "Review the error before running again."
        )
        sys.exit(1)
