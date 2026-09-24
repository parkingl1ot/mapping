import traceback
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from mapping_list import *


SOURCE_FILE = Path(r"C:\Users\xlin075\Documents\mapping\Book.xlsx")
OUTPUT_DIR = SOURCE_FILE.parent
LOG_PATH = OUTPUT_DIR / "processing_error_log.txt"
CONFLICT_PATH = OUTPUT_DIR / "mapping_conflicts.txt"


def log_error(message: str, exc: Exception | None = None) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    text = f"[{timestamp}] {message}\n"
    if exc is not None:
        text += traceback.format_exc() + "\n"
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(text)

def normalise_text(s): 
    if s is None:
        return ""
    return ''.join(ch for ch in str(s).upper() if ch.isalnum())

def contains_keyword(value, category) -> bool:
    if not value:
        return False
    norm_value = normalise_text(value)
    for keyword in category:
        if keyword in norm_value:
            return True
    return False


def make_copy_path() -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return OUTPUT_DIR / f"Book_copy_{timestamp}.xlsx"


def get_header_lookup(ws):
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20), values_only=True):
        cell_count = sum(1 for v in row if v is not None)
        if cell_count>11:
            return {str(v).strip(): idx + 1 for idx, v in enumerate(row) if v is not None}
    return {}


def process_workbook(source_path: Path, copy_path: Path):
    print("[1/7] Loading workbook from:", source_path)
    wb = load_workbook(source_path, data_only=False)
    print("[2/7] Workbook loaded. Sheets:", wb.sheetnames)

    print("[3/7] Creating workbook copy at:", copy_path)
    wb.save(copy_path)
    copied_wb = load_workbook(copy_path, data_only=False)

    errors = []
    conflicts = []
    total_rows_processed = 0
    total_updates = 0

    for sheet in copied_wb.worksheets:
        print(f"[4/7] Processing sheet: {sheet.title}")
        header_lookup = get_header_lookup(sheet)
        summary_index = header_lookup.get("摘要mapping")
        account_index = header_lookup.get("户名mapping")
        summary_triggers = [9, 11, 16, 17]
        account_triggers = [header_lookup.get("对方户名")]

        if "摘要mapping" not in header_lookup:
            log_error(f"  - No '摘要mapping' header found in {sheet.title}, please check the format of the book", Exception)
            print(f"  - No '摘要mapping' header found in {sheet.title}")
            continue

        if "户名mapping" not in header_lookup:
            log_error(f"  - No '户名mapping' header found in {sheet.title}, please check the format of the book", Exception)
            print(f"  - No '户名mapping' header found in {sheet.title}")
            continue

        max_row = sheet.max_row
        print(f"  - Data rows checked: {max_row - 1}")

        for category in summary.keys():
            for row_idx in range(2, max_row + 1):
                total_rows_processed += 1
                try:
                    ex_value = sheet.cell(row=row_idx, column=summary_index).value
                    trigger_hit = False

                    for col_idx in summary_triggers:
                        cell_value = sheet.cell(row=row_idx, column=col_idx).value
                        if contains_keyword(cell_value, summary.get(category)):
                            trigger_hit = True
                            break

                    if not trigger_hit:
                        continue

                    if ex_value:
                        conflicts.append({
                            "sheet": sheet.title,
                            "row": row_idx,
                            "existing_value": ex_value,
                            "proposed_value": category,
                            "trigger_columns": [str(get_column_letter(c) for c in summary_triggers)],
                        })
                        print(
                            f"  - 摘要conflict at sheet={sheet.title}, row={row_idx}: "
                            f"existing content='{ex_value}', proposed '{category}'."
                        )
                        continue

                    sheet.cell(row=row_idx, column=summary_index, value=category)
                    total_updates += 1
                    print(f"  - 摘要Updated sheet={sheet.title}, row={row_idx} => {category}")
                except Exception as exc:
                    errors.append((sheet.title, row_idx, str(exc)))
                    log_error(f"摘要Error while processing sheet={sheet.title}, row={row_idx}", exc)
                    print(f"  - 摘要Error on sheet={sheet.title}, row={row_idx}: {exc}")

        for category in account.keys():
            for row_idx in range(2, max_row + 1):
                total_rows_processed += 1
                try:
                    ex_value = sheet.cell(row=row_idx, column=account_index).value
                    trigger_hit = False

                    for col_idx in account_triggers:
                        cell_value = sheet.cell(row=row_idx, column=col_idx).value
                        if contains_keyword(cell_value, account.get(category)):
                            trigger_hit = True
                            break

                    if not trigger_hit:
                        continue

                    if ex_value:
                        conflicts.append({
                            "sheet": sheet.title,
                            "row": row_idx,
                            "existing_value": ex_value,
                            "proposed_value": category,
                        })
                        print(
                            f"  - 户名conflict at sheet={sheet.title}, row={row_idx}: "
                            f"existing content='{ex_value}', proposed '{category}'."
                        )
                        continue

                    sheet.cell(row=row_idx, column=account_index, value=category)
                    total_updates += 1
                    print(f"  - 户名Updated sheet={sheet.title}, row={row_idx} => {category}")
                except Exception as exc:
                    errors.append((sheet.title, row_idx, str(exc)))
                    log_error(f"户名Error while processing sheet={sheet.title}, row={row_idx}", exc)
                    print(f"  - 户名Error on sheet={sheet.title}, row={row_idx}: {exc}")


    print("[5/7] Saving updated workbook copy...")
    copied_wb.save(copy_path)

    if conflicts:
        with CONFLICT_PATH.open("w", encoding="utf-8") as f:
            f.write("Conflict report\n")
            f.write("================\n")
            for item in conflicts:
                f.write(
                    f"Sheet={item['sheet']}, Row={item['row']}, "
                    f"Existing content={item['existing_value']}, Proposed={item['proposed_value']}\n"
                )
        print(f"[6/7] {len(conflicts)} conflict(s) saved to: {CONFLICT_PATH}")
    else:
        CONFLICT_PATH.write_text("No conflicts detected.\n", encoding="utf-8")
        print("[6/7] No conflicts detected.")

    if errors:
        print(f"[7/7] Finished with {len(errors)} error(s). Please review {LOG_PATH}.")
    else:
        print(f"[7/7] Finished successfully. Updated {total_updates} row(s). Processed {total_rows_processed} row(s).")

    if conflicts:
        print("NOTICE: There are column J conflicts. The original value has been preserved and the user must manually handle these rows.")


def main():
    copy_path = make_copy_path()
    try:
        process_workbook(SOURCE_FILE, copy_path)
    except Exception as exc:
        log_error("Fatal script error.", exc)
        print("FATAL ERROR: The program encountered a critical error. See processing_error_log.txt for details.")
        raise


if __name__ == "__main__":
    main()
